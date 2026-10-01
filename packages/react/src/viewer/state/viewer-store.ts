import type { RangeAddress, UndoEntry, Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	CalcEngine,
	ConditionalFormatEngine,
	createProtectionGuard,
	EditTransaction,
	ProtectedError,
	StyleResolver,
} from 'xlsx-viewer-core';

import { autoFitTouchedRows } from '../../engine/layout/row-height';
import { singleCellSelection, type SelectionState } from '../../engine/selection/selection';
import { createTranslator, type Translate } from '../../i18n/translator';
import { initialSelection, needsCalculation } from './store-helpers';
import type { StoreSnapshot, UiState } from './types';
import { UsedExtents } from './used-extents';

export interface ViewerStoreOptions {
	readOnly?: boolean;
	/** Translator for text built outside React (default English); see `ViewerStore.t`. */
	t?: Translate;
}

export interface StoreCallbacks {
	onDirtyChange?: (dirty: boolean) => void;
	onChange?: () => void;
	onActiveSheetChange?: (index: number) => void;
	onSelectionChange?: (sheet: number, selection: SelectionState) => void;
}

interface HistoryItem {
	entry: UndoEntry;
	sheet: number;
	selectionBefore?: SelectionState;
	selectionAfter?: SelectionState;
}

const MAX_HISTORY = 200;

/**
 * Owns the workbook, its calculation engine, undo history and all viewer UI
 * state. React components subscribe with `useSyncExternalStore` through the
 * hooks in `use-store.ts`; mutations go through `transact`.
 */
export class ViewerStore {
	private readonly engine: CalcEngine;
	private calcBuilt = false;
	private calcTimer: ReturnType<typeof setTimeout> | undefined;
	readOnly: boolean;
	/**
	 * Translates UI text built outside React: status-bar messages, alerts and
	 * undo labels. ExcelViewer points it at the active dictionary; English by default.
	 */
	t: Translate;
	callbacks: StoreCallbacks = {};
	private listeners = new Set<() => void>();
	private ui: UiState;
	private docVersion = 0;
	private dirty = false;
	private undoStack: HistoryItem[] = [];
	private redoStack: HistoryItem[] = [];
	private snapshot: StoreSnapshot;
	private resolver?: StyleResolver;
	private resolverSig = '';
	private cfEngine?: ConditionalFormatEngine;

	constructor(
		readonly workbook: Workbook,
		options: ViewerStoreOptions = {},
	) {
		this.readOnly = options.readOnly ?? false;
		this.t = options.t ?? createTranslator();
		this.engine = new CalcEngine(workbook);
		if (needsCalculation(workbook)) {
			this.ensureCalc();
		} else {
			// Saved values are trusted: build the dependency graph after the first paint.
			this.calcTimer = setTimeout(() => this.ensureCalc(), 0);
		}
		const visible = workbook.sheets.findIndex((s) => s.state === 'visible');
		const active =
			workbook.sheets[workbook.view.activeTab]?.state === 'visible'
				? workbook.view.activeTab
				: Math.max(0, visible);
		const selections: Record<string, SelectionState> = {};
		for (const s of workbook.sheets) {
			selections[s.id] = initialSelection(s);
		}
		this.ui = {
			activeSheet: active,
			selections,
			edit: null,
			clipboard: null,
			scroll: {},
			dialog: null,
			ribbonTab: 'home',
			backstage: false,
			message: null,
			selectedObject: null,
			formatPainter: null,
			manualCalc: workbook.calc.calcMode === 'manual',
			formulaBarExpanded: false,
			traceArrows: [],
		};
		this.snapshot = this.makeSnapshot();
	}

	// ── Subscription ────────────────────────────────────────────────────

	subscribe = (listener: () => void): (() => void) => {
		this.listeners.add(listener);
		return () => this.listeners.delete(listener);
	};

	getSnapshot = (): StoreSnapshot => this.snapshot;

	private makeSnapshot(): StoreSnapshot {
		return {
			ui: this.ui,
			docVersion: this.docVersion,
			history: {
				undo: this.undoStack[this.undoStack.length - 1]?.entry.label ?? null,
				redo: this.redoStack[this.redoStack.length - 1]?.entry.label ?? null,
			},
			dirty: this.dirty,
		};
	}

	private emit(): void {
		this.snapshot = this.makeSnapshot();
		for (const l of this.listeners) {
			l();
		}
	}

	// ── UI state ────────────────────────────────────────────────────────

	get state(): UiState {
		return this.ui;
	}

	get version(): number {
		return this.docVersion;
	}

	get isDirty(): boolean {
		return this.dirty;
	}

	setUi(patch: Partial<UiState> | ((ui: UiState) => Partial<UiState>)): void {
		const p = typeof patch === 'function' ? patch(this.ui) : patch;
		const prevSheet = this.ui.activeSheet;
		this.ui = { ...this.ui, ...p };
		if (p.activeSheet !== undefined && p.activeSheet !== prevSheet) {
			this.workbook.view.activeTab = p.activeSheet;
			this.callbacks.onActiveSheetChange?.(p.activeSheet);
		}
		this.emit();
	}

	get activeSheet(): Worksheet {
		return this.workbook.sheets[this.ui.activeSheet];
	}

	selectionOf(sheetIndex = this.ui.activeSheet): SelectionState {
		const sheet = this.workbook.sheets[sheetIndex];
		return this.ui.selections[sheet.id] ?? singleCellSelection(0, 0);
	}

	setSelection(sel: SelectionState, sheetIndex = this.ui.activeSheet): void {
		const sheet = this.workbook.sheets[sheetIndex];
		this.ui = {
			...this.ui,
			selections: { ...this.ui.selections, [sheet.id]: sel },
			selectedObject: null,
		};
		sheet.view = {
			...sheet.view,
			selection: { activeCell: sel.active, ranges: sel.ranges },
		};
		this.callbacks.onSelectionChange?.(sheetIndex, sel);
		this.emit();
	}

	/** Shared style resolver; recreated when the stylesheet grows. */
	get styles(): StyleResolver {
		const s = this.workbook.styles;
		const sig = `${s.cellXfs.length}|${s.fonts.length}|${s.fills.length}|${s.borders.length}|${Object.keys(s.numFmts).length}|${s.dxfs.length}`;
		if (!this.resolver || sig !== this.resolverSig) {
			this.resolver = new StyleResolver(s, this.workbook.theme);
			this.resolverSig = sig;
		}
		return this.resolver;
	}

	/** Conditional-formatting evaluator (formula rules use the calc engine). */
	get conditionalFormats(): ConditionalFormatEngine {
		this.cfEngine ??= new ConditionalFormatEngine(this.workbook, (formula, sheet, row, col) =>
			this.calc.evaluateFormula(formula, sheet, row, col),
		);
		return this.cfEngine;
	}

	// ── Mutations ───────────────────────────────────────────────────────

	/**
	 * Run an editing operation as one undoable step. `fn` receives a
	 * transaction to `touch*` before mutating. Returns false when read-only.
	 */
	transact(
		label: string,
		fn: (tx: EditTransaction) => void,
		options: { selectionAfter?: SelectionState; autoRowHeight?: boolean } = {},
	): boolean {
		if (this.readOnly) {
			return false;
		}
		const sheet = this.ui.activeSheet;
		const selectionBefore = this.selectionOf(sheet);
		const tx = new EditTransaction(this.workbook, label);
		tx.guard = createProtectionGuard(this.workbook);
		try {
			fn(tx);
		} catch (e) {
			// Undo whatever the operation changed before it failed, then report the error.
			tx.commit().undo();
			this.docVersion++;
			this.cfEngine?.invalidate();
			if (e instanceof ProtectedError) {
				// Like Excel: an alert, and nothing changes.
				this.ui = { ...this.ui, dialog: { kind: 'alert', payload: { message: e.message } } };
				this.emit();
				return false;
			}
			this.emit();
			throw e;
		}
		if (options.autoRowHeight !== false) {
			// Automatic row heights follow content even on protected sheets (not a user format change).
			const guard = tx.guard;
			tx.guard = undefined;
			autoFitTouchedRows(tx, this.workbook, this.styles);
			tx.guard = guard;
		}
		const entry = tx.commit();
		this.recalcAfter(entry);
		if (this.ui.traceArrows.length) {
			this.ui = { ...this.ui, traceArrows: [] };
		}
		this.undoStack.push({ entry, sheet, selectionBefore, selectionAfter: options.selectionAfter });
		if (this.undoStack.length > MAX_HISTORY) {
			this.undoStack.shift();
		}
		this.redoStack = [];
		this.markChanged();
		if (options.selectionAfter) {
			this.setSelection(options.selectionAfter, Math.min(sheet, this.workbook.sheets.length - 1));
		} else {
			this.emit();
		}
		return true;
	}

	private readonly extents = new UsedExtents();

	/** Rows and columns spanned by content (scroll extents); see UsedExtents. */
	usedExtentOf(sheet: Worksheet): { rows: number; cols: number } {
		return this.extents.of(sheet);
	}

	private recalcAfter(entry: UndoEntry): void {
		this.extents.track(entry, this.workbook);
		if (this.ui.manualCalc) {
			return;
		}
		if (entry.structural || entry.fullRecalc) {
			this.calc.rebuild();
			this.calc.recalculateAll();
		} else if (entry.changedCells.length) {
			this.calc.onCellsChanged(entry.changedCells);
		}
	}

	/** Recalculate everything (F9). */
	recalculate(): void {
		this.calc.rebuild();
		this.calc.recalculateAll();
		this.cfEngine?.invalidate();
		this.docVersion++;
		this.emit();
	}

	private markChanged(): void {
		this.docVersion++;
		this.cfEngine?.invalidate();
		if (!this.dirty) {
			this.dirty = true;
			this.callbacks.onDirtyChange?.(true);
		}
		this.callbacks.onChange?.();
	}

	/** Notify that the workbook was mutated outside `transact` (not undoable). */
	touch(): void {
		this.extents.reset();
		this.markChanged();
		this.emit();
	}

	markSaved(): void {
		if (this.dirty) {
			this.dirty = false;
			this.callbacks.onDirtyChange?.(false);
			this.emit();
		}
	}

	undo(): void {
		const item = this.undoStack.pop();
		if (!item || this.readOnly) {
			return;
		}
		item.entry.undo();
		this.recalcAfter(item.entry);
		this.redoStack.push(item);
		this.markChanged();
		const sheet = Math.min(item.sheet, this.workbook.sheets.length - 1);
		if (sheet !== this.ui.activeSheet) {
			this.ui = { ...this.ui, activeSheet: sheet };
		}
		if (item.selectionBefore) {
			this.setSelection(item.selectionBefore, sheet);
		} else {
			this.emit();
		}
	}

	redo(): void {
		const item = this.redoStack.pop();
		if (!item || this.readOnly) {
			return;
		}
		item.entry.redo();
		this.recalcAfter(item.entry);
		this.undoStack.push(item);
		this.markChanged();
		const sheet = Math.min(item.sheet, this.workbook.sheets.length - 1);
		if (sheet !== this.ui.activeSheet) {
			this.ui = { ...this.ui, activeSheet: sheet };
		}
		this.setSelection(item.selectionAfter ?? this.selectionOf(sheet), sheet);
	}

	/** Ensure selection entries exist for sheets added since construction. */
	syncSheets(): void {
		let changed = false;
		const selections = { ...this.ui.selections };
		for (const s of this.workbook.sheets) {
			if (!selections[s.id]) {
				selections[s.id] = initialSelection(s);
				changed = true;
			}
		}
		let active = this.ui.activeSheet;
		if (active >= this.workbook.sheets.length) {
			active = this.workbook.sheets.length - 1;
			changed = true;
		}
		if (changed) {
			this.ui = { ...this.ui, selections, activeSheet: active };
		}
	}

	/** Ranges currently selected on the active sheet. */
	get selectedRanges(): RangeAddress[] {
		return this.selectionOf().ranges;
	}

	destroy(): void {
		this.listeners.clear();
		clearTimeout(this.calcTimer);
	}

	/** The calculation engine; its dependency graph is built on first use. */
	get calc(): CalcEngine {
		this.ensureCalc();
		return this.engine;
	}

	private ensureCalc(): void {
		if (!this.calcBuilt) {
			this.calcBuilt = true;
			clearTimeout(this.calcTimer);
			this.engine.rebuild();
		}
	}
}

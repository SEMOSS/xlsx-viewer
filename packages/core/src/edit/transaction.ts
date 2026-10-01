import { cellKey } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { RowInfo, Worksheet } from '../types/sheet';
import type { DefinedName, Workbook, WorkbookProtection, WorkbookView } from '../types/workbook';

/** A cell position within the workbook. */
export interface CellPosition {
	sheet: number;
	row: number;
	col: number;
}

export { restoreSheet, snapshotSheet, type SheetSnapshot } from './snapshot';
import { restoreSheet, snapshotSheet, type SheetSnapshot } from './snapshot';

interface WorkbookSnapshot {
	sheets: Worksheet[];
	definedNames: DefinedName[];
	view: WorkbookView;
	protection?: WorkbookProtection;
	stylesLengths: {
		fonts: number;
		fills: number;
		borders: number;
		xfs: number;
		dxfs: number;
		numFmts: string;
	};
}

/**
 * Why a sheet is touched, so a guard (sheet protection) can allow or block
 * the change: `edit` is any other change to the sheet, `references` the
 * formula adjustments other sheets receive, `sheets` sheet-level workbook
 * changes (rename, tab color, visibility).
 */
export type SheetChange =
	| 'edit'
	| 'structure'
	| 'insertRows'
	| 'insertColumns'
	| 'deleteRows'
	| 'deleteColumns'
	| 'formatColumns'
	| 'formatRows'
	| 'filter'
	| 'sort'
	| 'hyperlinks'
	| 'objects'
	| 'view'
	| 'references'
	| 'sheets';

/** Cell-level changes: values and formulas, or formatting only. */
export type CellChange = 'content' | 'format';

/** Workbook-level changes: the sheet list (add, delete, move, hide, rename) or names. */
export type WorkbookChange = 'sheets' | 'names';

/** Checks changes before they happen; throws to block them (e.g. ProtectedError). */
export interface TransactionGuard {
	cells?(sheetIndex: number, range: RangeAddress, change: CellChange): void;
	sheet?(sheetIndex: number, change: SheetChange): void;
	workbook?(change: WorkbookChange): void;
}

/** Undo/redo record produced by a committed transaction. */
export interface UndoEntry {
	label: string;
	undo(): void;
	redo(): void;
	/** Cells whose value or formula changed (for recalculation), per direction. */
	changedCells: CellPosition[];
	/** The transaction changed sheet structure (rows/cols/sheets), requiring a dependency rebuild. */
	structural: boolean;
	/** Too many cells changed to list: recalculate everything. */
	fullRecalc: boolean;
	/** Sheets whose cached layout must be refreshed. */
	sheets: number[];
}

/**
 * Records the state touched by an editing operation so it can be undone.
 * Call the `touch*` methods BEFORE mutating; `commit()` captures the
 * after-state and returns the undo entry.
 */
export class EditTransaction {
	private readonly cellBefore = new Map<Worksheet, Map<number, Cell | undefined>>();
	private readonly rowBefore = new Map<Worksheet, Map<number, RowInfo | undefined>>();
	private readonly sheetBefore = new Map<Worksheet, SheetSnapshot>();
	private workbookBefore?: WorkbookSnapshot;
	private readonly changed: CellPosition[] = [];
	private structural = false;
	private fullRecalc = false;
	private readonly touchedSheets = new Set<number>();

	/** Optional checks run by every touch (the viewer installs sheet protection here). */
	guard?: TransactionGuard;

	constructor(
		readonly workbook: Workbook,
		readonly label: string,
	) {}

	/** Remember cells (by range) before they change; also marks them as value-changed. */
	touchRange(sheetIndex: number, range: RangeAddress, valueChange = true): void {
		this.guard?.cells?.(sheetIndex, range, valueChange ? 'content' : 'format');
		const sheet = this.workbook.sheets[sheetIndex];
		const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
		this.touchedSheets.add(sheetIndex);
		if (area > 200_000 || this.sheetBefore.has(sheet)) {
			// Large areas: snapshot the sheet instead of individual cells.
			this.touchSheet(sheetIndex, false);
			if (valueChange) {
				if (area > 200_000) {
					this.fullRecalc = true;
				} else {
					for (let row = range.top; row <= range.bottom; row++) {
						for (let col = range.left; col <= range.right; col++) {
							this.changed.push({ sheet: sheetIndex, row, col });
						}
					}
				}
			}
			return;
		}
		let map = this.cellBefore.get(sheet);
		if (!map) {
			map = new Map();
			this.cellBefore.set(sheet, map);
		}
		for (let row = range.top; row <= range.bottom; row++) {
			for (let col = range.left; col <= range.right; col++) {
				const key = cellKey(row, col);
				if (!map.has(key)) {
					map.set(key, sheet.cells.get(key));
				}
				if (valueChange) {
					this.changed.push({ sheet: sheetIndex, row, col });
				}
			}
		}
	}

	touchCell(sheetIndex: number, row: number, col: number, valueChange = true): void {
		this.touchRange(sheetIndex, { top: row, left: col, bottom: row, right: col }, valueChange);
	}

	/**
	 * Remember row records (`sheet.rows`: height, hidden, row style) before
	 * they change. Much cheaper than `touchSheet` for a few rows.
	 */
	touchRows(sheetIndex: number, top: number, bottom: number): void {
		this.guard?.sheet?.(sheetIndex, 'formatRows');
		const sheet = this.workbook.sheets[sheetIndex];
		this.touchedSheets.add(sheetIndex);
		if (bottom - top > 100_000) {
			this.touchSheet(sheetIndex, false, 'formatRows');
			return;
		}
		let map = this.rowBefore.get(sheet);
		if (!map) {
			map = new Map();
			this.rowBefore.set(sheet, map);
		}
		for (let row = top; row <= bottom; row++) {
			if (!map.has(row)) {
				map.set(row, sheet.rows.get(row));
			}
		}
	}

	/** Keys (`cellKey`) of the cells of a sheet snapshotted one by one so far. */
	touchedCellKeys(sheetIndex: number): Iterable<number> {
		return this.cellBefore.get(this.workbook.sheets[sheetIndex])?.keys() ?? [];
	}

	/** Indices of the sheets touched so far. */
	touchedSheetIndices(): number[] {
		return [...this.touchedSheets];
	}

	/** Snapshot a whole sheet (rows, cols, merges, features...). */
	touchSheet(
		sheetIndex: number,
		structural = false,
		change: SheetChange = structural ? 'structure' : 'edit',
	): void {
		this.guard?.sheet?.(sheetIndex, change);
		this.touchedSheets.add(sheetIndex);
		const sheet = this.workbook.sheets[sheetIndex];
		if (!this.sheetBefore.has(sheet)) {
			this.sheetBefore.set(sheet, snapshotSheet(sheet));
		}
		if (structural) {
			this.structural = true;
		}
	}

	/** Snapshot the sheet list, names and view (sheet add / delete / reorder / rename). */
	touchWorkbook(change: WorkbookChange = 'names'): void {
		this.guard?.workbook?.(change);
		this.structural = true;
		if (!this.workbookBefore) {
			const s = this.workbook.styles;
			this.workbookBefore = {
				sheets: [...this.workbook.sheets],
				definedNames: this.workbook.definedNames.map((n) => ({ ...n })),
				view: { ...this.workbook.view },
				protection: this.workbook.protection,
				stylesLengths: {
					fonts: s.fonts.length,
					fills: s.fills.length,
					borders: s.borders.length,
					xfs: s.cellXfs.length,
					dxfs: s.dxfs.length,
					numFmts: JSON.stringify(s.numFmts),
				},
			};
		}
	}

	/** Run the sheet guard without taking a snapshot (e.g. sorting, which only touches cells). */
	checkSheet(sheetIndex: number, change: SheetChange): void {
		this.guard?.sheet?.(sheetIndex, change);
	}

	/** Run the workbook guard without taking a snapshot (e.g. tab color under structure protection). */
	checkWorkbook(change: WorkbookChange): void {
		this.guard?.workbook?.(change);
	}

	/** Note value changes for cells already snapshotted another way. */
	markChanged(pos: CellPosition): void {
		this.changed.push(pos);
	}

	commit(): UndoEntry {
		const wb = this.workbook;
		const cellAfter = new Map<Worksheet, Map<number, Cell | undefined>>();
		for (const [sheet, before] of this.cellBefore) {
			const after = new Map<number, Cell | undefined>();
			for (const key of before.keys()) {
				after.set(key, sheet.cells.get(key));
			}
			cellAfter.set(sheet, after);
			sheet.dirty = true;
		}
		const rowAfter = new Map<Worksheet, Map<number, RowInfo | undefined>>();
		for (const [sheet, before] of this.rowBefore) {
			const after = new Map<number, RowInfo | undefined>();
			for (const row of before.keys()) {
				after.set(row, sheet.rows.get(row));
			}
			rowAfter.set(sheet, after);
			sheet.dirty = true;
		}
		const sheetAfter = new Map<Worksheet, SheetSnapshot>();
		for (const sheet of this.sheetBefore.keys()) {
			sheet.dirty = true;
			sheetAfter.set(sheet, snapshotSheet(sheet));
		}
		const workbookBefore = this.workbookBefore;
		const workbookAfter: WorkbookSnapshot | undefined = workbookBefore
			? {
					sheets: [...wb.sheets],
					definedNames: wb.definedNames.map((n) => ({ ...n })),
					view: { ...wb.view },
					protection: wb.protection,
					stylesLengths: workbookBefore.stylesLengths,
				}
			: undefined;
		if (workbookBefore) {
			wb.dirty = true;
		}

		const apply = (
			cells: Map<Worksheet, Map<number, Cell | undefined>>,
			rows: Map<Worksheet, Map<number, RowInfo | undefined>>,
			sheets: Map<Worksheet, SheetSnapshot>,
			book: WorkbookSnapshot | undefined,
		): void => {
			if (book) {
				wb.sheets.splice(0, wb.sheets.length, ...book.sheets);
				wb.definedNames.splice(
					0,
					wb.definedNames.length,
					...book.definedNames.map((n) => ({ ...n })),
				);
				Object.assign(wb.view, book.view);
				wb.protection = book.protection;
				wb.dirty = true;
			}
			for (const snap of sheets.values()) {
				restoreSheet(snap);
				snap.sheet.dirty = true;
			}
			// Rows and cells after sheet snapshots: they may predate a later touchSheet.
			for (const [sheet, map] of rows) {
				for (const [row, info] of map) {
					if (info === undefined) {
						sheet.rows.delete(row);
					} else {
						sheet.rows.set(row, info);
					}
				}
				sheet.dirty = true;
			}
			for (const [sheet, map] of cells) {
				for (const [key, cell] of map) {
					if (cell === undefined) {
						sheet.cells.delete(key);
					} else {
						sheet.cells.set(key, cell);
					}
				}
				sheet.dirty = true;
			}
		};

		const changed = dedupePositions(this.changed);
		return {
			label: this.label,
			undo: () => apply(this.cellBefore, this.rowBefore, this.sheetBefore, workbookBefore),
			redo: () => apply(cellAfter, rowAfter, sheetAfter, workbookAfter),
			changedCells: changed,
			structural: this.structural,
			fullRecalc: this.fullRecalc,
			sheets: [...this.touchedSheets],
		};
	}
}

function dedupePositions(list: CellPosition[]): CellPosition[] {
	const seen = new Set<string>();
	const out: CellPosition[] = [];
	for (const p of list) {
		const k = `${p.sheet}:${p.row}:${p.col}`;
		if (!seen.has(k)) {
			seen.add(k);
			out.push(p);
		}
	}
	return out;
}

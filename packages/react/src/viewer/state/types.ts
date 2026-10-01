import type { RangeAddress } from 'xlsx-viewer-core';

import type { SelectionState } from '../../engine/selection/selection';

/** In-cell or formula-bar editing session. */
export interface EditState {
	/** Sheet index being edited. */
	sheet: number;
	row: number;
	col: number;
	/** Current editor text (formulas include the leading `=`). */
	text: string;
	/** Caret position within `text`. */
	caret: number;
	/**
	 * `enter`: started by typing over a cell; arrow keys commit and move.
	 * `edit`: started with F2 / double-click / the formula bar; arrows move the caret.
	 */
	mode: 'enter' | 'edit';
	/** Where the user is typing. */
	source: 'cell' | 'formulaBar';
	/** Text when editing started (to detect changes). */
	original: string;
	/** AutoComplete: the column entry `text` completes to (shown selected after the caret). */
	suggestion?: string;
	/**
	 * While typing a formula, the range being pointed at with the keyboard or
	 * mouse ("Point" mode) and where its reference text sits in `text`.
	 */
	point?: {
		range: RangeAddress;
		start: number;
		end: number;
		anchor: { row: number; col: number };
		sheet: number;
	};
}

export interface ClipboardMark {
	sheetId: string;
	ranges: RangeAddress[];
	cut: boolean;
}

export type RibbonTabId =
	| 'file'
	| 'home'
	| 'insert'
	| 'pageLayout'
	| 'formulas'
	| 'data'
	| 'review'
	| 'view'
	| 'help'
	| 'tableDesign'
	| 'chartDesign';

export interface DialogState {
	kind: string;
	payload?: unknown;
}

export type StatusMode = 'ready' | 'enter' | 'edit' | 'point';

export interface UiState {
	activeSheet: number;
	/** Selection per sheet, keyed by `Worksheet.id`. */
	selections: Record<string, SelectionState>;
	edit: EditState | null;
	clipboard: ClipboardMark | null;
	/** Scroll position per sheet id (restored when switching sheets). */
	scroll: Record<string, { left: number; top: number }>;
	dialog: DialogState | null;
	ribbonTab: RibbonTabId;
	/** File backstage open. */
	backstage: boolean;
	/** Transient status message shown in the status bar. */
	message: string | null;
	/** Selected floating object (picture / chart / shape) on the active sheet. */
	selectedObject: string | null;
	/** Format painter armed (`once` for single use, `sticky` after double-click). */
	formatPainter: { mode: 'once' | 'sticky'; sheet: number; range: RangeAddress } | null;
	/** Workbook-level calculation is manual (F9 to recalc). */
	manualCalc: boolean;
	/** Formula bar expanded to several lines. */
	formulaBarExpanded: boolean;
	/** Formula auditing arrows (Trace Precedents / Dependents); cleared by edits. */
	traceArrows: TraceArrow[];
}

/** One auditing arrow from a precedent area to a dependent cell, on `sheet`. */
export interface TraceArrow {
	sheet: number;
	from: RangeAddress;
	to: { row: number; col: number };
	kind: 'precedent' | 'dependent';
	/** The other end is on another sheet (drawn dashed from a sheet icon). */
	external?: boolean;
}

export interface StoreSnapshot {
	ui: UiState;
	/** Increments on every workbook mutation (content, styles, structure). */
	docVersion: number;
	/** Undo / redo availability and labels. */
	history: { undo: string | null; redo: string | null };
	dirty: boolean;
}

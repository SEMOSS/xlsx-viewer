import type { RangeAddress } from './address';
import type { FontSpec } from './style';

/** Excel error literals. */
export type ErrorCode =
	| '#NULL!'
	| '#DIV/0!'
	| '#VALUE!'
	| '#REF!'
	| '#NAME?'
	| '#NUM!'
	| '#N/A'
	| '#GETTING_DATA'
	| '#SPILL!'
	| '#CALC!'
	| '#FIELD!'
	| '#BLOCKED!'
	| '#CONNECT!'
	| '#BUSY!'
	| '#UNKNOWN!';

/**
 * An error value. Instances for each code are frozen singletons (see
 * `cellError()` in `values.ts`), so they can be compared by `code`.
 */
export interface CellError {
	readonly type: 'error';
	readonly code: ErrorCode;
}

/**
 * A cell's value: a number (dates and times are serial numbers), a string, a
 * boolean, an error, or `null` for blank.
 */
export type CellValue = number | string | boolean | CellError | null;

/** One run of a rich-text string. The concatenation of all run texts is the cell value. */
export interface RichTextRun {
	text: string;
	/** Run properties. Omitted properties inherit from the cell's font. */
	font?: FontSpec;
}

/** Array formula bookkeeping stored on the anchor (top-left) cell of an array result. */
export interface ArrayFormulaInfo {
	/** The full array range (legacy CSE block or the dynamic-array spill range). */
	ref: RangeAddress;
	/** `true` for an Excel 365 dynamic array formula (spills), `false` for a legacy CSE array. */
	dynamic: boolean;
}

/**
 * A cell in a worksheet.
 *
 * Treat Cell objects as immutable: editing operations replace the object in
 * `Worksheet.cells` rather than mutating it, which keeps undo snapshots cheap.
 */
export interface Cell {
	/** Literal value, or the cached / calculated result of `formula`. `null` = blank. */
	value: CellValue;
	/**
	 * Formula text in display form, without the leading `=`. Future-function
	 * prefixes (`_xlfn.`, `_xlws.`, `_xlpm.`) are stripped on load and restored on save.
	 * Shared formulas are expanded so every formula cell carries its own text.
	 */
	formula?: string;
	/** Present on the anchor cell of an array formula. */
	array?: ArrayFormulaInfo;
	/**
	 * Cell key (see `cellKey`) of the array anchor, on cells that hold part of
	 * another cell's array / spill result. Their `value` is owned by the anchor.
	 */
	arrayAnchor?: number;
	/** Index into `workbook.styles.cellXfs`. Omitted means 0 (the Normal style). */
	style?: number;
	/** Rich-text runs when the value is a formatted string. */
	richText?: RichTextRun[];
	/** OOXML cell-metadata index (`cm`), preserved for round-trip (dynamic arrays). */
	cm?: number;
	/** OOXML value-metadata index (`vm`), preserved for round-trip (rich data, images in cells). */
	vm?: number;
	/** Show phonetic run (`ph`). */
	ph?: boolean;
	/**
	 * The formula could not be recalculated (for example it uses a function the
	 * engine does not implement), so `value` is the last known cached result.
	 */
	stale?: boolean;
}

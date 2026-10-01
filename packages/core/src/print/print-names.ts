import { columnToLetters } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { formatRange } from '../address/range';
import { quoteSheetName } from '../address/sheet-name';
import { refAreaToRange, sameSheetName } from '../formula/language/references';
import { tokenizeFormula } from '../formula/language/tokenizer';
import type { RangeAddress } from '../types/address';
import type { DefinedName, Workbook } from '../types/workbook';
import type { PrintSpan } from './types';

/**
 * The built-in names behind Page Layout > Print Area and Print Titles:
 * sheet-scoped `_xlnm.Print_Area` (comma-separated areas) and
 * `_xlnm.Print_Titles` (whole columns and / or whole rows).
 */

export const PRINT_AREA_NAME = '_xlnm.Print_Area';
export const PRINT_TITLES_NAME = '_xlnm.Print_Titles';

function isBuiltin(def: DefinedName, local: string): boolean {
	const lower = def.name.toLowerCase();
	return lower === `_xlnm.${local.toLowerCase()}` || lower === local.toLowerCase();
}

/** The sheet's Print_Area / Print_Titles definition, if any. */
export function findPrintName(
	workbook: Workbook,
	sheetIndex: number,
	which: 'Print_Area' | 'Print_Titles',
): DefinedName | undefined {
	return workbook.definedNames.find((d) => d.localSheet === sheetIndex && isBuiltin(d, which));
}

/** References of a name formula that point at `sheetName` (or have no sheet). */
export function parseSheetReferences(formula: string, sheetName: string): RangeAddress[] {
	const out: RangeAddress[] = [];
	for (const t of tokenizeFormula(formula.trim().replace(/^=/, ''))) {
		if (t.type !== 'ref' || !t.area) {
			continue;
		}
		if (t.sheet && (t.sheet.workbook !== undefined || !sameSheetName(t.sheet.sheet, sheetName))) {
			continue;
		}
		out.push(refAreaToRange(t.area));
	}
	return out;
}

/** The print areas of a sheet (empty when it has none). */
export function getPrintArea(workbook: Workbook, sheetIndex: number): RangeAddress[] {
	const def = findPrintName(workbook, sheetIndex, 'Print_Area');
	const sheet = workbook.sheets[sheetIndex];
	return def && sheet ? parseSheetReferences(def.formula, sheet.name) : [];
}

export interface PrintTitles {
	/** Rows to repeat at top. */
	rows?: PrintSpan;
	/** Columns to repeat at left. */
	cols?: PrintSpan;
}

/** Rows to repeat at top and columns to repeat at left. */
export function getPrintTitles(workbook: Workbook, sheetIndex: number): PrintTitles {
	const def = findPrintName(workbook, sheetIndex, 'Print_Titles');
	const sheet = workbook.sheets[sheetIndex];
	const out: PrintTitles = {};
	if (!def || !sheet) {
		return out;
	}
	for (const r of parseSheetReferences(def.formula, sheet.name)) {
		const wholeRows = r.left === 0 && r.right === MAX_COL_INDEX;
		const wholeCols = r.top === 0 && r.bottom === MAX_ROW_INDEX;
		if (wholeRows && !wholeCols) {
			out.rows ??= { start: r.top, end: r.bottom };
		} else if (wholeCols && !wholeRows) {
			out.cols ??= { start: r.left, end: r.right };
		}
	}
	return out;
}

/** `Sheet1!$A$1:$D$20,Sheet1!$F$1:$G$5` for a list of areas. */
export function formatPrintArea(sheetName: string, areas: readonly RangeAddress[]): string {
	const prefix = `${quoteSheetName(sheetName)}!`;
	return areas.map((a) => prefix + formatRange(a, { absolute: true })).join(',');
}

/** `Sheet1!$A:$B,Sheet1!$1:$2` (columns first, as Excel writes them); '' when empty. */
export function formatPrintTitles(sheetName: string, titles: PrintTitles): string {
	const prefix = `${quoteSheetName(sheetName)}!`;
	const parts: string[] = [];
	if (titles.cols) {
		const { start, end } = titles.cols;
		parts.push(`${prefix}$${columnToLetters(start)}:$${columnToLetters(end)}`);
	}
	if (titles.rows) {
		parts.push(`${prefix}$${titles.rows.start + 1}:$${titles.rows.end + 1}`);
	}
	return parts.join(',');
}

/**
 * Parse a rows-to-repeat entry such as `$1:$3`, `1:3`, `$2` or `Sheet1!$1:$1`
 * (or a cell range, whose rows are used). `null` when invalid, `undefined`
 * when empty.
 */
export function parseTitleRows(text: string): PrintSpan | null | undefined {
	return parseTitleSpan(text, 'rows');
}

/** Parse a columns-to-repeat entry such as `$A:$B`, `A:A` or `$C`. */
export function parseTitleCols(text: string): PrintSpan | null | undefined {
	return parseTitleSpan(text, 'cols');
}

function parseTitleSpan(text: string, axis: 'rows' | 'cols'): PrintSpan | null | undefined {
	const raw = text.trim().replace(/^=/, '');
	if (raw === '') {
		return undefined;
	}
	const bang = raw.lastIndexOf('!');
	const ref = (bang >= 0 ? raw.slice(bang + 1) : raw).replace(/\$/g, '').trim();
	const single = axis === 'rows' ? /^\d{1,7}$/ : /^[A-Za-z]{1,3}$/;
	const expanded = single.test(ref) ? `${ref}:${ref}` : ref;
	const tokens = tokenizeFormula(expanded).filter((t) => t.type !== 'whitespace');
	const area = tokens.length === 1 && tokens[0].type === 'ref' ? tokens[0].area : undefined;
	if (!area || area.kind === (axis === 'rows' ? 'columns' : 'rows')) {
		return null;
	}
	const r = refAreaToRange(area);
	return axis === 'rows' ? { start: r.top, end: r.bottom } : { start: r.left, end: r.right };
}

/**
 * Parse a print-area entry such as `A1:D20`, `$A$1:$D$20,$F$1:$G$5` or
 * `Sheet1!A1:B2` (commas separate areas). `null` when it is not a list of
 * references; an empty list when the text is empty.
 */
export function parsePrintAreaText(text: string, sheetName: string): RangeAddress[] | null {
	const raw = text.trim().replace(/^=/, '');
	if (raw === '') {
		return [];
	}
	const out: RangeAddress[] = [];
	for (const t of tokenizeFormula(raw)) {
		if (t.type === 'whitespace' || t.type === 'union' || t.type === 'arg-sep') {
			continue;
		}
		if (t.type !== 'ref' || !t.area) {
			return null;
		}
		if (t.sheet && !sameSheetName(t.sheet.sheet, sheetName)) {
			return null;
		}
		out.push(refAreaToRange(t.area));
	}
	return out;
}

import { cellKey } from '../address/cell-address';
import { MAX_ROWS } from '../address/constants';
import { formatRange, isWholeColumnRange, isWholeRowRange } from '../address/range';
import { quoteSheetName } from '../address/sheet-name';
import { formulaResultToScalar, parseLiteralFormula } from '../conditional/formula-eval';
import { parseReferenceText, refAreaToRange } from '../formula/language/references';
import { tokenizeFormula } from '../formula/language/tokenizer';
import type { FormulaValue, Scalar } from '../formula/types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../formula/types';
import { formatValue } from '../numfmt/format-value';
import { findSheetIndex } from '../sheet/workbook';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { RangeAddress } from '../types/address';
import type { DefinedName, Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import { sanitizeDefinedName } from './names-syntax';

/**
 * Text helpers for the Define Name and Name Manager dialogs: the
 * "Refers to" text of a selection, the suggested name and the "Value"
 * column.
 */

/** Evaluates a formula (no leading `=`) as if entered in a cell (the calc engine's signature). */
export type DefinedNameEvaluator = (
	formula: string,
	sheetIndex: number,
	row: number,
	col: number,
) => FormulaValue;

/** Cells shown in the Value column of a multi-cell name before it is cut off with `...`. */
const MAX_VALUE_CELLS = 50;

function absoluteRangeText(r: RangeAddress): string {
	if (isWholeColumnRange(r) && isWholeRowRange(r)) {
		return `$1:$${MAX_ROWS}`;
	}
	return formatRange(r, { absolute: true });
}

/**
 * Absolute, sheet-qualified reference text of a selection, as Excel puts
 * it in "Refers to" (without the leading `=`): `Sheet1!$A$1:$B$5`,
 * `'My Sheet'!$C$3`, `Sheet1!$A:$A`, `Sheet1!$2:$2`; several ranges are
 * joined with `,`.
 */
export function selectionReference(sheetName: string, ranges: readonly RangeAddress[]): string {
	const prefix = `${quoteSheetName(sheetName)}!`;
	return ranges.map((r) => prefix + absoluteRangeText(r)).join(',');
}

function textLabel(workbook: Workbook, sheetIndex: number, row: number, col: number): string {
	if (row < 0 || col < 0) {
		return '';
	}
	const value = workbook.sheets[sheetIndex]?.cells.get(cellKey(row, col))?.value;
	return typeof value === 'string' ? value.trim() : '';
}

/**
 * The name Excel proposes in the New Name dialog: the text of the
 * selection's top-left cell when it holds a label, else the label just
 * above it, else the one to its left, made into a valid name
 * (`Unit Price` -> `Unit_Price`, `2024` -> `_2024`). `''` when none.
 */
export function suggestDefinedName(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
): string {
	const candidates: Array<[number, number]> = [
		[range.top, range.left],
		[range.top - 1, range.left],
		[range.top, range.left - 1],
	];
	for (const [row, col] of candidates) {
		const label = textLabel(workbook, sheetIndex, row, col);
		if (label !== '') {
			return sanitizeDefinedName(label);
		}
	}
	return '';
}

/** Displayed text of a cell (its value through its number format). */
function cellText(workbook: Workbook, sheetIndex: number, row: number, col: number): string {
	const cell = workbook.sheets[sheetIndex]?.cells.get(cellKey(row, col));
	if (!cell || cell.value === null) {
		return '';
	}
	const styles = workbook.styles;
	const code = numberFormatCode(styles.cellXfs[cell.style ?? 0]?.numFmtId ?? 0, styles.numFmts);
	return formatValue(cell.value, code, { date1904: workbook.date1904 }).text;
}

function quoted(text: string): string {
	return `"${text.replace(/"/g, '""')}"`;
}

/** Excel's array display: `{"a","b";"1","2"}`, cut off after a few dozen cells. */
function arrayText(rows: number, cols: number, item: (r: number, c: number) => string): string {
	const parts: string[] = [];
	let count = 0;
	for (let r = 0; r < rows && count < MAX_VALUE_CELLS; r++) {
		const items: string[] = [];
		for (let c = 0; c < cols && count < MAX_VALUE_CELLS; c++, count++) {
			items.push(item(r, c));
		}
		parts.push(items.join(','));
	}
	const truncated = rows * cols > MAX_VALUE_CELLS;
	return `{${parts.join(';')}${truncated ? '...' : ''}}`;
}

function rangeText(workbook: Workbook, sheetIndex: number, range: RangeAddress): string {
	const rows = range.bottom - range.top + 1;
	const cols = range.right - range.left + 1;
	if (rows === 1 && cols === 1) {
		return cellText(workbook, sheetIndex, range.top, range.left);
	}
	return arrayText(rows, cols, (r, c) =>
		quoted(cellText(workbook, sheetIndex, range.top + r, range.left + c)),
	);
}

function scalarText(value: Scalar, date1904: boolean): string {
	if (value === null) {
		return '';
	}
	if (isCellError(value)) {
		return value.code;
	}
	return formatValue(value, 'General', { date1904 }).text;
}

function resultText(workbook: Workbook, result: FormulaValue): string {
	if (isRefValue(result)) {
		return rangeText(workbook, result.sheet, result.range);
	}
	if (isMultiRefValue(result)) {
		const first = result.refs[0];
		return first ? rangeText(workbook, first.sheet, first.range) : '#VALUE!';
	}
	if (isArrayValue(result)) {
		if (result.rows === 1 && result.cols === 1) {
			return scalarText(formulaResultToScalar(result), workbook.date1904);
		}
		return arrayText(result.rows, result.cols, (r, c) => {
			const v = result.values[r]?.[c] ?? null;
			const text = scalarText(v, workbook.date1904);
			return typeof v === 'string' ? quoted(text) : text;
		});
	}
	return scalarText(formulaResultToScalar(result), workbook.date1904);
}

function hasRefError(formula: string): boolean {
	return tokenizeFormula(formula).some((t) => {
		if (t.type !== 'error') {
			return false;
		}
		const q = t.qualifierLength ?? 0;
		return t.text.slice(q, q + 5).toUpperCase() === '#REF!';
	});
}

/**
 * The Name Manager "Value" column: the displayed value of a one-cell
 * reference, `{"a","b";"1","2"}` for a multi-cell reference, `#REF!` for
 * a broken one, and otherwise the evaluated formula (constants are shown
 * even without an evaluator). `context` is the cell formulas are
 * evaluated at (default: A1 of the name's sheet or of the active sheet).
 */
export function definedNameDisplayValue(
	workbook: Workbook,
	def: DefinedName,
	evaluate?: DefinedNameEvaluator,
	context?: { sheet: number; row: number; col: number },
): string {
	const formula = def.formula.trim().replace(/^=/, '').trim();
	if (formula === '') {
		return '';
	}
	const fallbackSheet = Math.max(0, Math.min(workbook.view.activeTab, workbook.sheets.length - 1));
	const sheet = def.localSheet ?? context?.sheet ?? fallbackSheet;
	const ref = parseReferenceText(formula);
	if (ref && ref.sheet?.workbook === undefined && ref.sheet?.sheetEnd === undefined) {
		const index = ref.sheet ? findSheetIndex(workbook, ref.sheet.sheet) : sheet;
		return index < 0 ? '#REF!' : rangeText(workbook, index, refAreaToRange(ref.area));
	}
	if (hasRefError(formula)) {
		return '#REF!';
	}
	const result = evaluate?.(formula, sheet, context?.row ?? 0, context?.col ?? 0);
	if (result !== undefined && result !== null) {
		return resultText(workbook, result);
	}
	const literal = parseLiteralFormula(formula);
	return literal === undefined ? '' : scalarText(literal, workbook.date1904);
}

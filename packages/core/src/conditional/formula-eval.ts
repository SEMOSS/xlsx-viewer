import { cellKey } from '../address/cell-address';
import { translateFormula } from '../formula/translate';
import type { FormulaValue, Scalar } from '../formula/types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../formula/types';
import type { Workbook } from '../types/workbook';
import { cellError, ERRORS, parseErrorCode } from '../values/errors';
import { parseNumericText } from '../values/parse-number';

/**
 * Evaluates a formula (display form, without the leading `=`) as if it were
 * entered in the given cell, without storing it. Same signature as
 * `CalcEngine.evaluateFormula`, so the calc engine can be injected directly.
 */
export type FormulaEvaluator = (
	formula: string,
	sheetIndex: number,
	row: number,
	col: number,
) => FormulaValue;

/** Reads the current value of a cell (blank is `null`). */
export type CellReader = (sheetIndex: number, row: number, col: number) => Scalar;

/** A {@link CellReader} over the stored cell values of a workbook. */
export function workbookCellReader(workbook: Workbook): CellReader {
	return (sheetIndex, row, col) =>
		workbook.sheets[sheetIndex]?.cells.get(cellKey(row, col))?.value ?? null;
}

/**
 * Reduce a formula result to a single scalar. Conditional formatting and
 * validation formulas behave like array formulas whose first value counts:
 * references read their top-left cell and arrays yield their top-left element.
 */
export function formulaResultToScalar(result: FormulaValue | undefined, read?: CellReader): Scalar {
	if (result === undefined || result === null) {
		return null;
	}
	if (typeof result === 'number') {
		return Number.isFinite(result) ? result : ERRORS.NUM;
	}
	if (typeof result !== 'object') {
		return result;
	}
	if (isArrayValue(result)) {
		return formulaResultToScalar(result.values[0]?.[0] ?? null);
	}
	if (isRefValue(result)) {
		return read ? read(result.sheet, result.range.top, result.range.left) : ERRORS.VALUE;
	}
	if (isMultiRefValue(result)) {
		const first = result.refs[0];
		return first && read ? read(first.sheet, first.range.top, first.range.left) : ERRORS.VALUE;
	}
	if (isLambdaValue(result)) {
		return ERRORS.CALC;
	}
	return result;
}

/** Condition semantics: TRUE or a non-zero number. Text, blanks and errors are false. */
export function isConditionTrue(value: Scalar): boolean {
	return value === true || (typeof value === 'number' && value !== 0);
}

/**
 * Numeric value of a threshold result: numbers, numeric text and blank (0).
 * Booleans and errors are not usable as thresholds.
 */
export function scalarToNumber(value: Scalar, date1904 = false): number | undefined {
	if (typeof value === 'number') {
		return Number.isFinite(value) ? value : undefined;
	}
	if (value === null) {
		return 0;
	}
	if (typeof value === 'string') {
		return parseNumericText(value, { date1904 })?.value;
	}
	return undefined;
}

const NUMBER_LITERAL = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Whether text is a plain numeric literal such as `5`, `-1.5` or `2E3`. */
export function isNumericLiteral(text: string): boolean {
	return NUMBER_LITERAL.test(text.trim());
}

/**
 * Value of a formula that is a single constant: a number, a string literal
 * (`"abc"`), TRUE / FALSE or an error literal. `undefined` for anything that
 * needs the evaluator. Lets constant criteria skip formula evaluation.
 */
export function parseLiteralFormula(text: string): Scalar | undefined {
	const t = text.trim();
	if (t === '') {
		return undefined;
	}
	if (NUMBER_LITERAL.test(t)) {
		const n = Number(t);
		return Number.isFinite(n) ? n : undefined;
	}
	if (t.length >= 2 && t.startsWith('"') && t.endsWith('"')) {
		const inner = t.slice(1, -1);
		return /^(?:[^"]|"")*$/.test(inner) ? inner.replace(/""/g, '"') : undefined;
	}
	const upper = t.toUpperCase();
	if (upper === 'TRUE') {
		return true;
	}
	if (upper === 'FALSE') {
		return false;
	}
	const code = parseErrorCode(t);
	return code ? cellError(code) : undefined;
}

/** Remove one leading `=` (formulas typed by users). */
export function stripLeadingEquals(formula: string): string {
	const t = formula.trim();
	return t.startsWith('=') ? t.slice(1).trim() : t;
}

const ROW_OFFSET = 1 << 21;
const COL_SPAN = 1 << 16;
const COL_OFFSET = 1 << 15;
const MAX_CACHED_TRANSLATIONS = 4096;

/**
 * A formula written relative to an anchor cell (the top-left cell of the first
 * range of a conditional format or data validation). Evaluating it for
 * another cell shifts its relative references by the offset from the anchor.
 * Translations are cached per offset; constant formulas never hit the evaluator.
 */
export class RelativeFormula {
	/** Formula text without a leading `=`. */
	readonly text: string;
	/** Value of a constant formula, or `undefined` when it must be evaluated. */
	readonly literal: Scalar | undefined;
	private readonly relative: boolean;
	private readonly translations = new Map<number, string>();

	constructor(
		formula: string,
		readonly anchorRow: number,
		readonly anchorCol: number,
	) {
		this.text = stripLeadingEquals(formula);
		this.literal = parseLiteralFormula(this.text);
		this.relative = this.literal === undefined && translateFormula(this.text, 1, 1) !== this.text;
	}

	/** Formula text as it applies to cell (row, col). */
	textAt(row: number, col: number): string {
		const dr = row - this.anchorRow;
		const dc = col - this.anchorCol;
		if (!this.relative || (dr === 0 && dc === 0)) {
			return this.text;
		}
		const key = (dr + ROW_OFFSET) * COL_SPAN + (dc + COL_OFFSET);
		let text = this.translations.get(key);
		if (text === undefined) {
			if (this.translations.size >= MAX_CACHED_TRANSLATIONS) {
				this.translations.clear();
			}
			text = translateFormula(this.text, dr, dc);
			this.translations.set(key, text);
		}
		return text;
	}

	/** Raw evaluator result for cell (row, col); evaluator exceptions become `#VALUE!`. */
	evaluateRaw(
		evaluate: FormulaEvaluator,
		sheetIndex: number,
		row: number,
		col: number,
	): FormulaValue {
		if (this.literal !== undefined) {
			return this.literal;
		}
		try {
			return evaluate(this.textAt(row, col), sheetIndex, row, col) ?? null;
		} catch {
			return ERRORS.VALUE;
		}
	}

	/** Scalar result for cell (row, col) (see {@link formulaResultToScalar}). */
	evaluate(
		evaluate: FormulaEvaluator,
		sheetIndex: number,
		row: number,
		col: number,
		read?: CellReader,
	): Scalar {
		if (this.literal !== undefined) {
			return this.literal;
		}
		return formulaResultToScalar(this.evaluateRaw(evaluate, sheetIndex, row, col), read);
	}
}

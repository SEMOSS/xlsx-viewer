import { cellKey } from '../address/cell-address';
import { isConditionTrue, RelativeFormula, workbookCellReader } from '../conditional/formula-eval';
import { cellValueText } from '../conditional/text-match';
import type { Scalar } from '../formula/types';
import type { Cell, CellValue } from '../types/cell';
import type { DataValidation, DataValidationOperator } from '../types/data-validation';
import { parseNumericText } from '../values/parse-number';
import { listEntryMatches, resolveListSource } from './list-source';
import {
	DEFAULT_VALIDATION_ERROR_MESSAGE,
	DEFAULT_VALIDATION_ERROR_TITLE,
	type DataValidationContext,
	type DataValidationResult,
} from './types';

type Operand =
	| { kind: 'number'; value: number }
	| { kind: 'blank' }
	| { kind: 'missing' }
	| { kind: 'invalid' };

const MISSING: Operand = { kind: 'missing' };
const INVALID: Operand = { kind: 'invalid' };

/** Evaluate a criteria formula relative to the validated cell (anchor: top-left of the first range). */
function evaluateCriterion(
	formula: string | undefined,
	dv: DataValidation,
	ctx: DataValidationContext,
): Scalar | undefined {
	if (formula === undefined || formula.trim() === '') {
		return undefined;
	}
	const anchor = dv.ranges[0];
	const relative = new RelativeFormula(formula, anchor?.top ?? ctx.row, anchor?.left ?? ctx.col);
	const read = ctx.workbook ? workbookCellReader(ctx.workbook) : undefined;
	return relative.evaluate(ctx.evaluate, ctx.sheetIndex, ctx.row, ctx.col, read);
}

function toNumberOrUndefined(value: Scalar, date1904: boolean): number | undefined {
	if (typeof value === 'number') {
		return Number.isFinite(value) ? value : undefined;
	}
	if (typeof value === 'string') {
		return parseNumericText(value, { date1904 })?.value;
	}
	return undefined;
}

function numericOperand(
	formula: string | undefined,
	dv: DataValidation,
	ctx: DataValidationContext,
): Operand {
	const v = evaluateCriterion(formula, dv, ctx);
	if (v === undefined) {
		return MISSING;
	}
	if (v === null) {
		return { kind: 'blank' };
	}
	const n = toNumberOrUndefined(v, ctx.date1904);
	return n === undefined ? INVALID : { kind: 'number', value: n };
}

/** Compare at Excel's 15 significant digits so computed bounds match typed values. */
function norm(n: number): number {
	return Number(n.toPrecision(15));
}

function applyOperator(op: DataValidationOperator, value: number, a: number, b: number): boolean {
	const v = norm(value);
	const x = norm(a);
	const y = norm(b);
	switch (op) {
		case 'between':
			return v >= Math.min(x, y) && v <= Math.max(x, y);
		case 'notBetween':
			return v < Math.min(x, y) || v > Math.max(x, y);
		case 'equal':
			return v === x;
		case 'notEqual':
			return v !== x;
		case 'greaterThan':
			return v > x;
		case 'greaterThanOrEqual':
			return v >= x;
		case 'lessThan':
			return v < x;
		case 'lessThanOrEqual':
			return v <= x;
		default:
			return true;
	}
}

/**
 * Check a number against the rule's operator and criteria. Like Excel, when
 * "Ignore blank" is on and a criterion refers to a blank cell, any value is
 * accepted; without it the blank counts as 0. Missing criteria impose nothing.
 */
function checkCriteria(value: number, dv: DataValidation, ctx: DataValidationContext): boolean {
	const op = dv.operator ?? 'between';
	const two = op === 'between' || op === 'notBetween';
	const a = numericOperand(dv.formula1, dv, ctx);
	const b = two ? numericOperand(dv.formula2, dv, ctx) : MISSING;
	if ((a.kind === 'blank' || b.kind === 'blank') && dv.allowBlank) {
		return true;
	}
	if (a.kind === 'invalid' || b.kind === 'invalid') {
		return false;
	}
	if (a.kind === 'missing' || (two && b.kind === 'missing')) {
		return true;
	}
	const x = a.kind === 'number' ? a.value : 0;
	const y = b.kind === 'number' ? b.value : 0;
	return applyOperator(op, value, x, y);
}

/**
 * Run `fn` with `value` temporarily stored in the validated cell, so custom
 * formulas see the value being entered (Excel evaluates them that way).
 */
function withPendingValue<T>(ctx: DataValidationContext, value: CellValue, fn: () => T): T {
	const sheet = ctx.workbook?.sheets[ctx.sheetIndex];
	if (!sheet) {
		return fn();
	}
	const key = cellKey(ctx.row, ctx.col);
	const previous = sheet.cells.get(key);
	const pending: Cell = { value };
	if (previous?.style !== undefined) {
		pending.style = previous.style;
	}
	sheet.cells.set(key, pending);
	try {
		return fn();
	} finally {
		if (previous) {
			sheet.cells.set(key, previous);
		} else {
			sheet.cells.delete(key);
		}
	}
}

function checkValue(value: CellValue, dv: DataValidation, ctx: DataValidationContext): boolean {
	if (!dv.type || dv.type === 'none') {
		return true;
	}
	if (value === null || value === '') {
		if (dv.allowBlank) {
			return true;
		}
		if (dv.type !== 'custom' && dv.type !== 'textLength') {
			return false;
		}
	}
	switch (dv.type) {
		case 'whole':
		case 'decimal':
		case 'date':
		case 'time': {
			const n = typeof value === 'boolean' ? undefined : toNumberOrUndefined(value, ctx.date1904);
			if (n === undefined || (dv.type === 'whole' && !Number.isInteger(n))) {
				return false;
			}
			return checkCriteria(n, dv, ctx);
		}
		case 'textLength': {
			const text = cellValueText(value);
			return text !== undefined && checkCriteria(text.length, dv, ctx);
		}
		case 'list': {
			const source = resolveListSource(dv, ctx);
			if (!source) {
				return true;
			}
			// Documented Excel quirk: "Ignore blank" with a named source range that
			// contains blank cells accepts any value.
			if (dv.allowBlank && source.named && source.hasBlank) {
				return true;
			}
			const toNumber = (text: string) => parseNumericText(text, { date1904: ctx.date1904 })?.value;
			return source.entries.some((entry) => listEntryMatches(entry, value, toNumber));
		}
		case 'custom': {
			if (!dv.formula1) {
				return true;
			}
			const result = withPendingValue(ctx, value, () => evaluateCriterion(dv.formula1, dv, ctx));
			return isConditionTrue(result ?? null);
		}
		default:
			return true;
	}
}

/**
 * Validate a value about to be entered in a cell (pass the parsed value;
 * numeric text is also accepted for numeric types). Criteria formulas are
 * evaluated relative to the cell. Invalid results carry the alert to show
 * (the rule's style, title and message, or Excel's defaults) when the rule
 * enables the error alert (`showErrorMessage: true`, which Excel's dialog
 * sets by default). Without it the result has no style: like Excel, the
 * value is accepted silently (it is still reported as invalid).
 */
export function validateValue(
	value: CellValue,
	dv: DataValidation,
	ctx: DataValidationContext,
): DataValidationResult {
	if (checkValue(value, dv, ctx)) {
		return { valid: true };
	}
	if (dv.showErrorMessage !== true) {
		return { valid: false };
	}
	return {
		valid: false,
		style: dv.errorStyle ?? 'stop',
		title: dv.errorTitle || DEFAULT_VALIDATION_ERROR_TITLE,
		message: dv.error || DEFAULT_VALIDATION_ERROR_MESSAGE,
	};
}

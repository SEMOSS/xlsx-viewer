import { MAX_COLS, MAX_ROWS } from '../../../address/constants';
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import { etaDefinition } from '../../evaluator/lambda';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	LambdaValue,
	Scalar,
} from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from '../logical/lift';
import { clampToUsed } from '../reference/shared';

/**
 * Shared plumbing of the LAMBDA helper functions: argument sources that
 * hand out cells as references (so `ROW(x)` or `ISBLANK(x)` work inside the
 * LAMBDA) or array elements, and conversion of LAMBDA results to a single
 * value (a nested array is `#CALC!`).
 */

/** A grid of LAMBDA arguments. */
export interface Source {
	readonly rows: number;
	readonly cols: number;
	/** Element (r, c); single rows / columns broadcast, outside the grid is `#N/A`. */
	at(r: number, c: number): FormulaValue;
	/** Row r as a reference or array (BYROW). */
	row(r: number): FormulaValue;
	/** Column c as a reference or array (BYCOL). */
	column(c: number): FormulaValue;
}

function arraySource(a: ArrayValue): Source {
	return {
		rows: a.rows,
		cols: a.cols,
		at: (r, c) => {
			const rr = a.rows === 1 ? 0 : r;
			const cc = a.cols === 1 ? 0 : c;
			return rr < a.rows && cc < a.cols ? a.values[rr][cc] : ERRORS.NA;
		},
		row: (r) =>
			a.cols === 1
				? a.values[r][0]
				: { kind: 'array', rows: 1, cols: a.cols, values: [a.values[r]] },
		column: (c) =>
			a.rows === 1
				? a.values[0][c]
				: { kind: 'array', rows: a.rows, cols: 1, values: a.values.map((line) => [line[c]]) },
	};
}

/** View an argument as a grid of LAMBDA inputs (`#VALUE!` / errors for unusable arguments). */
export function toSource(arg: FunctionArg, ctx: FunctionContext): Source | CellError {
	const ref = isRefValue(arg)
		? arg
		: isMultiRefValue(arg) && arg.refs.length === 1
			? arg.refs[0]
			: null;
	if (ref !== null) {
		if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
			return ERRORS.VALUE;
		}
		const { top, left, bottom, right } = clampToUsed(ref, ctx);
		const sheet = ref.sheet;
		const rows = bottom - top + 1;
		const cols = right - left + 1;
		const cell = (r: number, c: number): FormulaValue => ({
			kind: 'ref',
			sheet,
			range: { top: top + r, left: left + c, bottom: top + r, right: left + c },
		});
		return {
			rows,
			cols,
			at: (r, c) => {
				const rr = rows === 1 ? 0 : r;
				const cc = cols === 1 ? 0 : c;
				return rr < rows && cc < cols ? cell(rr, cc) : ERRORS.NA;
			},
			row: (r) => ({ kind: 'ref', sheet, range: { top: top + r, left, bottom: top + r, right } }),
			column: (c) => ({
				kind: 'ref',
				sheet,
				range: { top, left: left + c, bottom, right: left + c },
			}),
		};
	}
	if (isMultiRefValue(arg) || isLambdaValue(arg) || isMissingNode(arg)) {
		return ERRORS.VALUE;
	}
	if (isArrayValue(arg)) {
		return arg.rows === 0 || arg.cols === 0 ? ERRORS.CALC : arraySource(arg);
	}
	return arraySource({ kind: 'array', rows: 1, cols: 1, values: [[arg as Scalar]] });
}

/** The LAMBDA argument of a helper (the last one), or an error. */
export function lambdaArg(arg: FunctionArg | undefined): LambdaValue | CellError {
	if (isLambdaValue(arg)) {
		return arg;
	}
	return isCellError(arg) ? arg : ERRORS.VALUE;
}

/** The LAMBDA can be called with `n` arguments. */
export function acceptsArgs(fn: LambdaValue, n: number): boolean {
	const eta = etaDefinition(fn);
	if (eta !== undefined) {
		return n >= eta.minArgs && n <= eta.maxArgs;
	}
	const required = fn.params.filter((p) => !p.startsWith('[')).length;
	return n >= required && n <= fn.params.length;
}

/** A LAMBDA result as one array element: single cells are read, nested arrays are `#CALC!`. */
export function elementResult(v: FormulaValue, ctx: FunctionContext): Scalar {
	if (isRefValue(v)) {
		const r = v.range;
		if (r.top === r.bottom && r.left === r.right && (v.sheetEnd ?? v.sheet) === v.sheet) {
			return ctx.getCellValue(v.sheet, r.top, r.left);
		}
		return ERRORS.CALC;
	}
	if (isArrayValue(v)) {
		return v.rows === 1 && v.cols === 1 ? v.values[0][0] : ERRORS.CALC;
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? elementResult(v.refs[0], ctx) : ERRORS.VALUE;
	}
	if (isLambdaValue(v)) {
		return ERRORS.CALC;
	}
	if (isMissingNode(v)) {
		return null;
	}
	return v as Scalar;
}

/** A positive dimension argument (MAKEARRAY rows / columns). */
export function dimensionArg(
	arg: FunctionArg,
	ctx: FunctionContext,
	max: number,
): number | CellError {
	const n = toNumber(isMissingNode(arg) ? null : (arg as Scalar), ctx.date1904);
	if (typeof n !== 'number') {
		return n;
	}
	const k = Math.trunc(n);
	return k < 1 || k > max ? ERRORS.VALUE : k;
}

export const MAX_DIMENSIONS = { rows: MAX_ROWS, cols: MAX_COLS } as const;

/** Build an array from a per-element function. */
export function fill(rows: number, cols: number, f: (r: number, c: number) => Scalar): ArrayValue {
	const values: Scalar[][] = new Array(rows);
	for (let r = 0; r < rows; r++) {
		const line: Scalar[] = new Array(cols);
		for (let c = 0; c < cols; c++) {
			line[c] = f(r, c);
		}
		values[r] = line;
	}
	return { kind: 'array', rows, cols, values };
}

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../../../address/constants';
import type { CellError } from '../../../../types/cell';
import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionArg, FunctionContext, RefValue, Scalar } from '../../../types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../../../types';
import { isMissing } from '../args';

/**
 * A range argument of the conditional aggregates, read by position. `rows`
 * and `cols` are the declared shape; `values` may be smaller because
 * references are clamped to the used area, and positions outside it read
 * as blank.
 */
export interface Grid {
	rows: number;
	cols: number;
	values: Scalar[][];
	filledRows: number;
	filledCols: number;
}

/** View a reference, array or scalar argument as a grid. */
export function toGrid(arg: FunctionArg, ctx: FunctionContext): Grid | CellError {
	if (isRefValue(arg)) {
		if (arg.sheetEnd !== undefined && arg.sheetEnd !== arg.sheet) {
			return ERRORS.VALUE;
		}
		const r = arg.range;
		const resolved = ctx.resolveRef(arg);
		return {
			rows: r.bottom - r.top + 1,
			cols: r.right - r.left + 1,
			values: resolved.values,
			filledRows: resolved.rows,
			filledCols: resolved.cols,
		};
	}
	if (isArrayValue(arg)) {
		return {
			rows: arg.rows,
			cols: arg.cols,
			values: arg.values,
			filledRows: arg.rows,
			filledCols: arg.cols,
		};
	}
	if (isMultiRefValue(arg)) {
		return ERRORS.VALUE;
	}
	if (isCellError(arg)) {
		return arg;
	}
	const v = isMissing(arg) ? null : (arg as Scalar);
	return { rows: 1, cols: 1, values: [[v]], filledRows: 1, filledCols: 1 };
}

/** Value at a position of the declared shape (blank outside the filled area). */
export function gridAt(g: Grid, r: number, c: number): Scalar {
	return r < g.filledRows && c < g.filledCols ? (g.values[r]?.[c] ?? null) : null;
}

export function sameShape(a: Grid, b: Grid): boolean {
	return a.rows === b.rows && a.cols === b.cols;
}

/**
 * The block of `ref` that starts at its top-left cell and has the given
 * shape (AVERAGEIF resizes its average range to match the criteria range).
 */
export function resizeRef(ref: RefValue, rows: number, cols: number): RefValue {
	const top = ref.range.top;
	const left = ref.range.left;
	return {
		kind: 'ref',
		sheet: ref.sheet,
		range: {
			top,
			left,
			bottom: Math.min(MAX_ROW_INDEX, top + rows - 1),
			right: Math.min(MAX_COL_INDEX, left + cols - 1),
		},
	};
}

/** Criteria argument as a scalar (omitted criteria behave as a blank cell). */
export function criteriaValue(arg: FunctionArg): Scalar {
	return isMissing(arg) ? null : (arg as Scalar);
}

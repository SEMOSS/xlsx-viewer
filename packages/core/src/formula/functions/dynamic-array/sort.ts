import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import { toArrayValue } from '../../invoke';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	Scalar,
} from '../../types';
import { arrayArg, isOmitted, optionalBoolean } from '../lookup/args';
import { columnValues, pickCols, pickRows, vectorValues } from './shared';
import { sortPermutation } from './sort-order';

/**
 * An optional array argument taken as `any`, so that an omitted argument
 * (`SORT(a,,-1)`) can be told apart from a reference to a blank cell.
 */
function optionalArray(
	args: readonly FunctionArg[],
	i: number,
	ctx: FunctionContext,
): ArrayValue | CellError | undefined {
	return isOmitted(args, i) ? undefined : toArrayValue(args[i] as FormulaValue, ctx);
}

/** sort_index values (1-based, scalar or 1-D array) as 0-based key positions within `width`. */
function readSortIndices(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	width: number,
): number[] | CellError {
	const a = optionalArray(args, 1, ctx);
	if (a === undefined) {
		return [0];
	}
	if (isCellError(a)) {
		return a;
	}
	const list = vectorValues(a);
	if (!list) {
		return ERRORS.VALUE;
	}
	const out: number[] = [];
	for (const v of list) {
		const n = toNumber(v, ctx.date1904);
		if (isCellError(n)) {
			return n;
		}
		const k = Math.trunc(n);
		if (k < 1 || k > width) {
			return ERRORS.VALUE;
		}
		out.push(k - 1);
	}
	return out;
}

/** One sort order value: 1 ascending, -1 descending, anything else `#VALUE!`. */
function readOrder(v: Scalar, date1904: boolean): boolean | CellError {
	const n = toNumber(v, date1904);
	if (isCellError(n)) {
		return n;
	}
	if (n === 1 || n === -1) {
		return n === -1;
	}
	return ERRORS.VALUE;
}

/** sort_order values, one per key (a single value applies to every key). */
function readSortOrders(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	keys: number,
): boolean[] | CellError {
	const a = optionalArray(args, 2, ctx);
	if (a === undefined) {
		return new Array<boolean>(keys).fill(false);
	}
	if (isCellError(a)) {
		return a;
	}
	const list = vectorValues(a);
	if (!list || (list.length !== 1 && list.length !== keys)) {
		return ERRORS.VALUE;
	}
	const out: boolean[] = [];
	for (const v of list) {
		const d = readOrder(v, ctx.date1904);
		if (isCellError(d)) {
			return d;
		}
		out.push(d);
	}
	return out.length === keys ? out : new Array<boolean>(keys).fill(out[0]);
}

/**
 * SORT(array, [sort_index], [sort_order], [by_col])
 *
 * Sorts rows (or columns with by_col TRUE) by one or more keys: sort_index
 * and sort_order may be arrays for a multi-level sort. The sort is stable
 * and follows Excel's order (numbers, text, logicals, errors; blanks last).
 */
export const SORT: FunctionDefinition = {
	name: 'SORT',
	minArgs: 1,
	maxArgs: 4,
	params: ['array', 'any', 'any', 'value'],
	filePrefix: '_xlfn._xlws.',
	returnsArray: true,
	impl: (args, ctx) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const byCol = optionalBoolean(args, 3, false);
		if (isCellError(byCol)) {
			return byCol;
		}
		const indices = readSortIndices(args, ctx, byCol ? a.rows : a.cols);
		if (isCellError(indices)) {
			return indices;
		}
		const orders = readSortOrders(args, ctx, indices.length);
		if (isCellError(orders)) {
			return orders;
		}
		const keys = indices.map((k) => (byCol ? a.values[k] : columnValues(a, k)));
		const perm = sortPermutation(byCol ? a.cols : a.rows, keys, orders);
		return byCol ? pickCols(a, perm) : pickRows(a, perm);
	},
};

/**
 * SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)
 *
 * Sorts rows by columns of keys (or columns by rows of keys). Every by_array
 * is a single column as tall as array, or a single row as wide as it, all in
 * the same orientation; otherwise `#VALUE!`. Each sort_order is 1 or -1.
 */
export const SORTBY: FunctionDefinition = {
	name: 'SORTBY',
	minArgs: 2,
	maxArgs: Infinity,
	params: ['array', 'any'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const bys: ArrayValue[] = [];
		const descending: boolean[] = [];
		let rowsFit = true;
		let colsFit = true;
		for (let i = 1; i < args.length; i += 2) {
			const by = optionalArray(args, i, ctx) ?? ERRORS.VALUE;
			if (isCellError(by)) {
				return by;
			}
			rowsFit &&= by.cols === 1 && by.rows === a.rows;
			colsFit &&= by.rows === 1 && by.cols === a.cols;
			if (!rowsFit && !colsFit) {
				return ERRORS.VALUE;
			}
			let desc = false;
			const o = optionalArray(args, i + 1, ctx);
			if (o !== undefined) {
				if (isCellError(o)) {
					return o;
				}
				if (o.rows !== 1 || o.cols !== 1) {
					return ERRORS.VALUE;
				}
				const d = readOrder(o.values[0][0], ctx.date1904);
				if (isCellError(d)) {
					return d;
				}
				desc = d;
			}
			bys.push(by);
			descending.push(desc);
		}
		const keys = bys.map((by) => vectorValues(by) as Scalar[]);
		const perm = sortPermutation(rowsFit ? a.rows : a.cols, keys, descending);
		return rowsFit ? pickRows(a, perm) : pickCols(a, perm);
	},
};

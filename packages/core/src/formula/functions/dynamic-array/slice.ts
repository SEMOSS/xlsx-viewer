import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type { FormulaValue, FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { arrayArg } from '../lookup/args';
import { buildArray, MAX_ARRAY_CELLS, optionalCount, padArg, pickCols, pickRows } from './shared';

/** Half-open index span kept by TAKE (`take`) or DROP along a dimension of length `len`. */
function span(len: number, n: number | undefined, take: boolean): [number, number] {
	if (n === undefined) {
		return [0, len];
	}
	if (take) {
		return n >= 0 ? [0, Math.min(n, len)] : [Math.max(0, len + n), len];
	}
	return n >= 0 ? [Math.min(n, len), len] : [0, Math.max(0, len + n)];
}

function takeOrDrop(args: FunctionArg[], ctx: FunctionContext, take: boolean): FormulaValue {
	const a = arrayArg(args, 0);
	if (isCellError(a)) {
		return a;
	}
	const rows = optionalCount(args, 1, ctx.date1904);
	if (isCellError(rows)) {
		return rows;
	}
	const cols = optionalCount(args, 2, ctx.date1904);
	if (isCellError(cols)) {
		return cols;
	}
	const [r0, r1] = span(a.rows, rows, take);
	const [c0, c1] = span(a.cols, cols, take);
	if (r1 <= r0 || c1 <= c0) {
		return ERRORS.CALC;
	}
	return buildArray(r1 - r0, c1 - c0, (r, c) => a.values[r0 + r][c0 + c]);
}

/**
 * TAKE(array, rows, [columns]): the first rows / columns (last when
 * negative); counts beyond the array keep everything, zero is `#CALC!`.
 */
export const TAKE: FunctionDefinition = {
	name: 'TAKE',
	minArgs: 2,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => takeOrDrop(args, ctx, true),
};

/**
 * DROP(array, rows, [columns]): without the first rows / columns (last when
 * negative); dropping everything is `#CALC!`.
 */
export const DROP: FunctionDefinition = {
	name: 'DROP',
	minArgs: 2,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => takeOrDrop(args, ctx, false),
};

function choose(args: FunctionArg[], ctx: FunctionContext, byCol: boolean): FormulaValue {
	const a = arrayArg(args, 0);
	if (isCellError(a)) {
		return a;
	}
	const len = byCol ? a.cols : a.rows;
	const picks: number[] = [];
	for (let i = 1; i < args.length; i++) {
		const sel = arrayArg(args, i);
		if (isCellError(sel)) {
			return sel;
		}
		for (const line of sel.values) {
			for (const v of line) {
				const n = toNumber(v, ctx.date1904);
				if (isCellError(n)) {
					return n;
				}
				const k = Math.trunc(n);
				if (k === 0 || Math.abs(k) > len) {
					return ERRORS.VALUE;
				}
				picks.push(k > 0 ? k - 1 : len + k);
			}
		}
	}
	if (picks.length * (byCol ? a.rows : a.cols) > MAX_ARRAY_CELLS) {
		return ERRORS.NUM;
	}
	return byCol ? pickCols(a, picks) : pickRows(a, picks);
}

/**
 * CHOOSECOLS(array, col_num1, [col_num2], ...): the listed columns in order
 * (negative numbers count from the right; arrays of numbers are allowed).
 * Zero or a number beyond the array is `#VALUE!`.
 */
export const CHOOSECOLS: FunctionDefinition = {
	name: 'CHOOSECOLS',
	minArgs: 2,
	maxArgs: Infinity,
	params: ['array'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => choose(args, ctx, true),
};

/** CHOOSEROWS(array, row_num1, [row_num2], ...): the listed rows, like CHOOSECOLS. */
export const CHOOSEROWS: FunctionDefinition = {
	name: 'CHOOSEROWS',
	minArgs: 2,
	maxArgs: Infinity,
	params: ['array'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => choose(args, ctx, false),
};

/**
 * EXPAND(array, rows, [columns], [pad_with])
 *
 * Grows the array to rows x columns (omitted sizes keep the current one),
 * padding with pad_with or `#N/A`. A size smaller than the array is
 * `#VALUE!`. Errors are allowed as the padding value.
 */
export const EXPAND: FunctionDefinition = {
	name: 'EXPAND',
	minArgs: 2,
	maxArgs: 4,
	params: ['array', 'value', 'value', 'value'],
	propagateErrors: false,
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const rows = optionalCount(args, 1, ctx.date1904) ?? a.rows;
		if (isCellError(rows)) {
			return rows;
		}
		const cols = optionalCount(args, 2, ctx.date1904) ?? a.cols;
		if (isCellError(cols)) {
			return cols;
		}
		if (rows < a.rows || cols < a.cols) {
			return ERRORS.VALUE;
		}
		if (rows * cols > MAX_ARRAY_CELLS) {
			return ERRORS.NUM;
		}
		const pad = padArg(args, 3);
		return buildArray(rows, cols, (r, c) => (r < a.rows && c < a.cols ? a.values[r][c] : pad));
	},
};

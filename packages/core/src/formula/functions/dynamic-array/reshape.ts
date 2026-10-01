import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	Scalar,
} from '../../types';
import { arrayArg, optionalBoolean, optionalInteger, scalarArg } from '../lookup/args';
import { buildArray, MAX_ARRAY_CELLS, padArg, vectorValues } from './shared';

function stack(args: FunctionArg[], vertical: boolean): FormulaValue {
	const parts: ArrayValue[] = [];
	let rows = 0;
	let cols = 0;
	for (let i = 0; i < args.length; i++) {
		const a = arrayArg(args, i);
		if (isCellError(a)) {
			return a;
		}
		parts.push(a);
		rows = vertical ? rows + a.rows : Math.max(rows, a.rows);
		cols = vertical ? Math.max(cols, a.cols) : cols + a.cols;
	}
	if (rows * cols > MAX_ARRAY_CELLS) {
		return ERRORS.NUM;
	}
	if (vertical) {
		const values: Scalar[][] = [];
		for (const a of parts) {
			for (const src of a.values) {
				const line: Scalar[] = new Array(cols);
				for (let c = 0; c < cols; c++) {
					line[c] = c < a.cols ? src[c] : ERRORS.NA;
				}
				values.push(line);
			}
		}
		return { kind: 'array', rows, cols, values };
	}
	const values: Scalar[][] = Array.from({ length: rows }, () => new Array<Scalar>(cols));
	let offset = 0;
	for (const a of parts) {
		for (let r = 0; r < rows; r++) {
			const src = r < a.rows ? a.values[r] : null;
			const line = values[r];
			for (let c = 0; c < a.cols; c++) {
				line[offset + c] = src ? src[c] : ERRORS.NA;
			}
		}
		offset += a.cols;
	}
	return { kind: 'array', rows, cols, values };
}

/** VSTACK(array1, ...): arrays appended vertically; narrower ones are padded with `#N/A`. */
export const VSTACK: FunctionDefinition = {
	name: 'VSTACK',
	minArgs: 1,
	maxArgs: Infinity,
	params: ['array'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args) => stack(args, true),
};

/** HSTACK(array1, ...): arrays appended horizontally; shorter ones are padded with `#N/A`. */
export const HSTACK: FunctionDefinition = {
	name: 'HSTACK',
	minArgs: 1,
	maxArgs: Infinity,
	params: ['array'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args) => stack(args, false),
};

function flatten(args: FunctionArg[], ctx: FunctionContext, toColumn: boolean): FormulaValue {
	const a = arrayArg(args, 0);
	if (isCellError(a)) {
		return a;
	}
	const ignore = optionalInteger(args, 1, 0, ctx.date1904);
	if (isCellError(ignore)) {
		return ignore;
	}
	if (ignore < 0 || ignore > 3) {
		return ERRORS.VALUE;
	}
	const byCol = optionalBoolean(args, 2, false);
	if (isCellError(byCol)) {
		return byCol;
	}
	const skipBlanks = (ignore & 1) !== 0;
	const skipErrors = (ignore & 2) !== 0;
	const out: Scalar[] = [];
	const outer = byCol ? a.cols : a.rows;
	const inner = byCol ? a.rows : a.cols;
	for (let i = 0; i < outer; i++) {
		for (let j = 0; j < inner; j++) {
			const v = byCol ? a.values[j][i] : a.values[i][j];
			if ((skipBlanks && v === null) || (skipErrors && isCellError(v))) {
				continue;
			}
			out.push(v);
		}
	}
	if (out.length === 0) {
		return ERRORS.CALC;
	}
	return toColumn
		? { kind: 'array', rows: out.length, cols: 1, values: out.map((v) => [v]) }
		: { kind: 'array', rows: 1, cols: out.length, values: [out] };
}

/**
 * TOCOL(array, [ignore], [scan_by_column]): all values in one column,
 * scanned by row (or by column). ignore: 0 keep all, 1 skip blanks, 2 skip
 * errors, 3 skip both. Nothing left is `#CALC!`.
 */
export const TOCOL: FunctionDefinition = {
	name: 'TOCOL',
	minArgs: 1,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => flatten(args, ctx, true),
};

/** TOROW(array, [ignore], [scan_by_column]): like TOCOL, into one row. */
export const TOROW: FunctionDefinition = {
	name: 'TOROW',
	minArgs: 1,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => flatten(args, ctx, false),
};

function wrap(args: FunctionArg[], ctx: FunctionContext, intoRows: boolean): FormulaValue {
	const a = arrayArg(args, 0);
	if (isCellError(a)) {
		return a;
	}
	const items = vectorValues(a);
	if (!items) {
		return ERRORS.VALUE;
	}
	const n = toNumber(scalarArg(args, 1), ctx.date1904);
	if (isCellError(n)) {
		return n;
	}
	const count = Math.trunc(n);
	if (count < 1) {
		return ERRORS.NUM;
	}
	const lines = Math.ceil(items.length / count);
	if (lines * count > MAX_ARRAY_CELLS) {
		return ERRORS.NUM;
	}
	const pad = padArg(args, 2);
	const at = (k: number): Scalar => (k < items.length ? items[k] : pad);
	return intoRows
		? buildArray(lines, count, (r, c) => at(r * count + c))
		: buildArray(count, lines, (r, c) => at(c * count + r));
}

/**
 * WRAPROWS(vector, wrap_count, [pad_with]): a row or column wrapped into
 * rows of wrap_count values, the last padded with pad_with or `#N/A`. A 2-D
 * input is `#VALUE!`, wrap_count below 1 `#NUM!`.
 */
export const WRAPROWS: FunctionDefinition = {
	name: 'WRAPROWS',
	minArgs: 2,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	propagateErrors: false,
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => wrap(args, ctx, true),
};

/** WRAPCOLS(vector, wrap_count, [pad_with]): like WRAPROWS, filling columns. */
export const WRAPCOLS: FunctionDefinition = {
	name: 'WRAPCOLS',
	minArgs: 2,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	propagateErrors: false,
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => wrap(args, ctx, false),
};

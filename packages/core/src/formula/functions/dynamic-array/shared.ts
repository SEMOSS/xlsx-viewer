import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type { ArrayValue, FunctionArg, Scalar } from '../../types';
import { isOmitted, scalarArg } from '../lookup/args';

/**
 * Largest array a dynamic-array function builds. Excel answers `#NUM!` when
 * a result is "too large"; this guard keeps a runaway SEQUENCE or EXPAND
 * from exhausting memory in the browser.
 */
export const MAX_ARRAY_CELLS = 10_000_000;

/** Validate result dimensions: negative is `#VALUE!`, empty is `#CALC!`, oversized is `#NUM!`. */
export function checkDimensions(rows: number, cols: number): CellError | null {
	if (rows < 0 || cols < 0) {
		return ERRORS.VALUE;
	}
	if (rows === 0 || cols === 0) {
		return ERRORS.CALC;
	}
	return rows * cols > MAX_ARRAY_CELLS ? ERRORS.NUM : null;
}

/** Build a `rows` x `cols` array from a cell function. */
export function buildArray(
	rows: number,
	cols: number,
	cell: (r: number, c: number) => Scalar,
): ArrayValue {
	const values: Scalar[][] = new Array(rows);
	for (let r = 0; r < rows; r++) {
		const line: Scalar[] = new Array(cols);
		for (let c = 0; c < cols; c++) {
			line[c] = cell(r, c);
		}
		values[r] = line;
	}
	return { kind: 'array', rows, cols, values };
}

/** The given rows of an array, in the given order. */
export function pickRows(a: ArrayValue, rows: readonly number[]): ArrayValue {
	const values: Scalar[][] = new Array(rows.length);
	for (let i = 0; i < rows.length; i++) {
		values[i] = a.values[rows[i]].slice();
	}
	return { kind: 'array', rows: rows.length, cols: a.cols, values };
}

/** The given columns of an array, in the given order. */
export function pickCols(a: ArrayValue, cols: readonly number[]): ArrayValue {
	const values: Scalar[][] = new Array(a.rows);
	for (let r = 0; r < a.rows; r++) {
		const src = a.values[r];
		const line: Scalar[] = new Array(cols.length);
		for (let i = 0; i < cols.length; i++) {
			line[i] = src[cols[i]];
		}
		values[r] = line;
	}
	return { kind: 'array', rows: a.rows, cols: cols.length, values };
}

/** Column `c` of an array as a list. */
export function columnValues(a: ArrayValue, c: number): Scalar[] {
	const out: Scalar[] = new Array(a.rows);
	for (let r = 0; r < a.rows; r++) {
		out[r] = a.values[r][c];
	}
	return out;
}

/** A 1-D array (single row or column) as a list; `null` for a 2-D array. */
export function vectorValues(a: ArrayValue): Scalar[] | null {
	if (a.rows === 1) {
		return a.values[0].slice();
	}
	return a.cols === 1 ? columnValues(a, 0) : null;
}

/**
 * Optional whole-number argument, truncated toward zero: `undefined` when
 * omitted, or the coercion error.
 */
export function optionalCount(
	args: readonly FunctionArg[],
	i: number,
	date1904: boolean,
): number | undefined | CellError {
	if (isOmitted(args, i)) {
		return undefined;
	}
	const n = toNumber(scalarArg(args, i), date1904);
	return typeof n === 'number' ? Math.trunc(n) : n;
}

/** Optional padding value: `#N/A` when omitted, otherwise the value itself (errors included). */
export function padArg(args: readonly FunctionArg[], i: number): Scalar {
	return isOmitted(args, i) ? ERRORS.NA : scalarArg(args, i);
}

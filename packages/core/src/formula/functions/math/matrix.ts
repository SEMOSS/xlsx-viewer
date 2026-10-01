import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { asArray } from '../../coerce';
import type { ArrayValue, FunctionDefinition, Scalar } from '../../types';
import { makeArray } from '../../types';
import { integerArg } from './args';
import { XLFN } from './define';
import { cleanNoise } from './rounding';

/** Largest identity matrix MUNIT builds (it must fit the grid's columns). */
const MAX_MUNIT = 16_384;
const MAX_MUNIT_CELLS = 4_000_000;

/**
 * Numeric matrix of an array argument. Errors propagate; text, booleans
 * and blanks are #VALUE! (matrix functions accept numbers only).
 */
function toMatrix(arr: ArrayValue): number[][] | CellError {
	const out: number[][] = [];
	for (const row of arr.values) {
		const line: number[] = [];
		for (const v of row) {
			if (isCellError(v)) {
				return v;
			}
			if (typeof v !== 'number') {
				return ERRORS.VALUE;
			}
			line.push(v);
		}
		out.push(line);
	}
	return out;
}

/** Determinant by LU decomposition with partial pivoting. */
export function determinant(m: number[][]): number {
	const n = m.length;
	const a = m.map((row) => row.slice());
	let det = 1;
	for (let col = 0; col < n; col++) {
		let pivot = col;
		for (let r = col + 1; r < n; r++) {
			if (Math.abs(a[r][col]) > Math.abs(a[pivot][col])) {
				pivot = r;
			}
		}
		if (a[pivot][col] === 0) {
			return 0;
		}
		if (pivot !== col) {
			[a[pivot], a[col]] = [a[col], a[pivot]];
			det = -det;
		}
		det *= a[col][col];
		for (let r = col + 1; r < n; r++) {
			const f = a[r][col] / a[col][col];
			for (let c = col; c < n; c++) {
				a[r][c] -= f * a[col][c];
			}
		}
	}
	return det;
}

/** Inverse by Gauss-Jordan elimination, or null when the matrix is singular. */
export function inverse(m: number[][]): number[][] | null {
	const n = m.length;
	const a = m.map((row, i) => [...row, ...row.map((_, j) => (i === j ? 1 : 0))]);
	let scale = 0;
	for (const row of m) {
		for (const v of row) {
			scale = Math.max(scale, Math.abs(v));
		}
	}
	const tolerance = scale * n * Number.EPSILON;
	for (let col = 0; col < n; col++) {
		let pivot = col;
		for (let r = col + 1; r < n; r++) {
			if (Math.abs(a[r][col]) > Math.abs(a[pivot][col])) {
				pivot = r;
			}
		}
		if (Math.abs(a[pivot][col]) <= tolerance) {
			return null;
		}
		[a[pivot], a[col]] = [a[col], a[pivot]];
		const p = a[col][col];
		for (let c = 0; c < 2 * n; c++) {
			a[col][c] /= p;
		}
		for (let r = 0; r < n; r++) {
			if (r !== col && a[r][col] !== 0) {
				const f = a[r][col];
				for (let c = 0; c < 2 * n; c++) {
					a[r][c] -= f * a[col][c];
				}
			}
		}
	}
	return a.map((row) => row.slice(n));
}

/** Tidy matrix results: entries that are pure round-off relative to the largest entry become 0. */
function tidy(rows: number[][]): Scalar[][] {
	let max = 0;
	for (const row of rows) {
		for (const v of row) {
			max = Math.max(max, Math.abs(v));
		}
	}
	return rows.map((row) => row.map((v) => (Math.abs(v) <= max * 1e-15 ? 0 : cleanNoise(v))));
}

const MDETERM: FunctionDefinition = {
	name: 'MDETERM',
	minArgs: 1,
	maxArgs: 1,
	params: ['array'],
	impl: (args) => {
		const arr = asArray(args[0]);
		if (arr.rows !== arr.cols || arr.rows === 0) {
			return ERRORS.VALUE;
		}
		const m = toMatrix(arr);
		if (isCellError(m)) {
			return m;
		}
		const det = determinant(m);
		if (!Number.isFinite(det)) {
			return ERRORS.NUM;
		}
		// The determinant of an integer matrix is an integer: drop the round-off.
		const integral = m.every((row) => row.every(Number.isInteger));
		const rounded = Math.round(det);
		if (integral && Math.abs(det - rounded) <= Math.max(1, Math.abs(det)) * 1e-10) {
			return rounded === 0 ? 0 : rounded;
		}
		return cleanNoise(det);
	},
};

const MINVERSE: FunctionDefinition = {
	name: 'MINVERSE',
	minArgs: 1,
	maxArgs: 1,
	params: ['array'],
	returnsArray: true,
	impl: (args) => {
		const arr = asArray(args[0]);
		if (arr.rows !== arr.cols || arr.rows === 0) {
			return ERRORS.VALUE;
		}
		const m = toMatrix(arr);
		if (isCellError(m)) {
			return m;
		}
		const inv = inverse(m);
		if (!inv || inv.some((row) => row.some((v) => !Number.isFinite(v)))) {
			return ERRORS.NUM;
		}
		return makeArray(tidy(inv));
	},
};

const MMULT: FunctionDefinition = {
	name: 'MMULT',
	minArgs: 2,
	maxArgs: 2,
	params: ['array', 'array'],
	returnsArray: true,
	impl: (args) => {
		const a = asArray(args[0]);
		const b = asArray(args[1]);
		if (a.cols !== b.rows || a.rows === 0 || b.cols === 0) {
			return ERRORS.VALUE;
		}
		const ma = toMatrix(a);
		if (isCellError(ma)) {
			return ma;
		}
		const mb = toMatrix(b);
		if (isCellError(mb)) {
			return mb;
		}
		const out: Scalar[][] = [];
		for (let r = 0; r < a.rows; r++) {
			const row: Scalar[] = [];
			for (let c = 0; c < b.cols; c++) {
				let sum = 0;
				for (let k = 0; k < a.cols; k++) {
					sum += ma[r][k] * mb[k][c];
				}
				if (!Number.isFinite(sum)) {
					return ERRORS.NUM;
				}
				row.push(sum);
			}
			out.push(row);
		}
		return makeArray(out);
	},
};

const MUNIT: FunctionDefinition = {
	name: 'MUNIT',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	filePrefix: XLFN,
	returnsArray: true,
	impl: (args, ctx) => {
		const n = integerArg(args[0], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		if (n < 1 || n > MAX_MUNIT) {
			return ERRORS.VALUE;
		}
		if (n * n > MAX_MUNIT_CELLS) {
			return ERRORS.NUM;
		}
		const rows: Scalar[][] = [];
		for (let r = 0; r < n; r++) {
			const row = new Array<Scalar>(n).fill(0);
			row[r] = 1;
			rows.push(row);
		}
		return makeArray(rows);
	},
};

export const MATRIX_FUNCTIONS: FunctionDefinition[] = [MDETERM, MINVERSE, MMULT, MUNIT];

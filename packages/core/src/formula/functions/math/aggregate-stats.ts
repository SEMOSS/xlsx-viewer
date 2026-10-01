import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import { approxCeil, snapToInteger } from './rounding';
import { sumOf } from './sums';

/**
 * Statistical kernels used by SUBTOTAL and AGGREGATE. They operate on the
 * numbers already collected from the arguments (the statistical function
 * module has its own argument handling for the stand-alone functions).
 */

type Result = number | CellError;

export function average(values: readonly number[]): Result {
	return values.length === 0 ? ERRORS.DIV0 : sumOf(values) / values.length;
}

export function variance(values: readonly number[], sample: boolean): Result {
	const n = values.length;
	if (n === 0 || (sample && n < 2)) {
		return ERRORS.DIV0;
	}
	const mean = sumOf(values) / n;
	const squares = sumOf(values.map((v) => (v - mean) * (v - mean)));
	return squares / (sample ? n - 1 : n);
}

export function stdev(values: readonly number[], sample: boolean): Result {
	const v = variance(values, sample);
	return typeof v === 'number' ? Math.sqrt(v) : v;
}

export function max(values: readonly number[]): number {
	let m = -Infinity;
	for (const v of values) {
		m = v > m ? v : m;
	}
	return values.length === 0 ? 0 : m;
}

export function min(values: readonly number[]): number {
	let m = Infinity;
	for (const v of values) {
		m = v < m ? v : m;
	}
	return values.length === 0 ? 0 : m;
}

export function product(values: readonly number[]): number {
	if (values.length === 0) {
		return 0;
	}
	let p = 1;
	for (const v of values) {
		p *= v;
	}
	return p;
}

function sorted(values: readonly number[]): number[] {
	return [...values].sort((a, b) => a - b);
}

export function median(values: readonly number[]): Result {
	const n = values.length;
	if (n === 0) {
		return ERRORS.NUM;
	}
	const s = sorted(values);
	const mid = Math.floor(n / 2);
	return n % 2 === 1 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** Most frequent value; ties go to the value that occurs first. #N/A when nothing repeats. */
export function modeSingle(values: readonly number[]): Result {
	const counts = new Map<number, number>();
	for (const v of values) {
		counts.set(v, (counts.get(v) ?? 0) + 1);
	}
	let best: number | undefined;
	let bestCount = 1;
	for (const v of values) {
		const c = counts.get(v)!;
		if (c > bestCount) {
			best = v;
			bestCount = c;
		}
	}
	return best === undefined ? ERRORS.NA : best;
}

/** LARGE / SMALL: `k` is rounded up to an integer and must be within 1..n. */
export function kth(values: readonly number[], k: number, largest: boolean): Result {
	const kk = approxCeil(k);
	if (values.length === 0 || kk < 1 || kk > values.length) {
		return ERRORS.NUM;
	}
	const s = sorted(values);
	return largest ? s[s.length - kk] : s[kk - 1];
}

function interpolate(s: readonly number[], h: number): number {
	const position = snapToInteger(h);
	const lo = Math.floor(position);
	const frac = position - lo;
	if (lo + 1 >= s.length || frac === 0) {
		return s[lo];
	}
	return s[lo] + frac * (s[lo + 1] - s[lo]);
}

export function percentileInc(values: readonly number[], k: number): Result {
	if (values.length === 0 || k < 0 || k > 1) {
		return ERRORS.NUM;
	}
	const s = sorted(values);
	return interpolate(s, (s.length - 1) * k);
}

export function percentileExc(values: readonly number[], k: number): Result {
	const n = values.length;
	if (n === 0 || k <= 0 || k >= 1) {
		return ERRORS.NUM;
	}
	const h = snapToInteger((n + 1) * k) - 1;
	if (h < 0 || h > n - 1) {
		return ERRORS.NUM;
	}
	return interpolate(sorted(values), h);
}

export function quartileInc(values: readonly number[], quart: number): Result {
	const q = Math.trunc(quart);
	return q < 0 || q > 4 ? ERRORS.NUM : percentileInc(values, q / 4);
}

export function quartileExc(values: readonly number[], quart: number): Result {
	const q = Math.trunc(quart);
	return q <= 0 || q >= 4 ? ERRORS.NUM : percentileExc(values, q / 4);
}

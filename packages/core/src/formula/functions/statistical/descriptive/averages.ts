import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { aggregateFunction, coerceArgs, define, finish, numbersOf, type NumResult } from '../args';
import { approxFloor, devsq, mean, sortedAscending, sum } from './core';

/** Options of the `*A` functions: TRUE counts 1, FALSE and text count 0 inside references. */
export const A_OPTIONS = { boolsInRefs: true, textInRefsAsZero: true } as const;

function average(xs: number[]): NumResult {
	return xs.length === 0 ? ERRORS.DIV0 : sum(xs) / xs.length;
}

function avedev(xs: number[]): NumResult {
	if (xs.length === 0) {
		return ERRORS.NUM;
	}
	const m = mean(xs);
	let acc = 0;
	for (const x of xs) {
		acc += Math.abs(x - m);
	}
	return acc / xs.length;
}

/**
 * Geometric mean from a running product kept in [1, 2) with the binary
 * exponent tracked separately: no overflow, and no log / exp round trip.
 */
function geomean(xs: number[]): NumResult {
	const n = xs.length;
	if (n === 0) {
		return ERRORS.NUM;
	}
	let mantissa = 1;
	let exponent = 0;
	for (const x of xs) {
		if (!(x > 0) || x === Infinity) {
			return ERRORS.NUM;
		}
		const e = Math.floor(Math.log2(x));
		mantissa *= x / Math.pow(2, e);
		exponent += e;
		const k = Math.floor(Math.log2(mantissa));
		mantissa /= Math.pow(2, k);
		exponent += k;
	}
	const q = Math.floor(exponent / n);
	const r = exponent - q * n;
	return Math.pow(mantissa, 1 / n) * Math.pow(2, r / n) * Math.pow(2, q);
}

function harmean(xs: number[]): NumResult {
	if (xs.length === 0) {
		return ERRORS.NA;
	}
	let acc = 0;
	for (const x of xs) {
		if (!(x > 0)) {
			return ERRORS.NUM;
		}
		acc += 1 / x;
	}
	return xs.length / acc;
}

/** Median of numbers (sorted copy). */
export function median(xs: readonly number[]): NumResult {
	const n = xs.length;
	if (n === 0) {
		return ERRORS.NUM;
	}
	const s = sortedAscending(xs);
	const mid = n >> 1;
	return n % 2 === 1 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

const TRIMMEAN: FunctionDefinition = define({
	name: 'TRIMMEAN',
	minArgs: 2,
	maxArgs: 2,
	params: ['any', 'value'],
	impl: (args, ctx) => {
		const xs = numbersOf([args[0]], ctx);
		if (isCellError(xs)) {
			return xs;
		}
		const v = coerceArgs(['n'], args, ctx, 1);
		if (isCellError(v)) {
			return v;
		}
		const percent = v[0];
		if (xs.length === 0 || percent < 0 || percent >= 1) {
			return ERRORS.NUM;
		}
		// Excel trims floor(n * percent / 2) points from each end.
		const k = approxFloor((xs.length * percent) / 2);
		const s = sortedAscending(xs);
		return finish(mean(s.slice(k, s.length - k)));
	},
});

export const AVERAGE_FUNCTIONS: FunctionDefinition[] = [
	aggregateFunction('AVERAGE', average),
	aggregateFunction('AVERAGEA', average, A_OPTIONS),
	aggregateFunction('AVEDEV', avedev),
	aggregateFunction('DEVSQ', (xs) => (xs.length === 0 ? ERRORS.NUM : devsq(xs))),
	aggregateFunction('GEOMEAN', geomean),
	aggregateFunction('HARMEAN', harmean),
	aggregateFunction('MEDIAN', median),
	TRIMMEAN,
];

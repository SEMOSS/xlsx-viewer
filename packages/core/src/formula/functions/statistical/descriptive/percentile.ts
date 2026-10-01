import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../../types';
import { type ArgSpec, coerceArgs, define, finish, numbersOf, type NumResult } from '../args';
import { approxFloor, sortedAscending } from './core';

/** Inclusive percentile of sorted data: interpolation at rank (n - 1) k. */
export function percentileInc(sorted: readonly number[], k: number): NumResult {
	const n = sorted.length;
	if (n === 0 || k < 0 || k > 1) {
		return ERRORS.NUM;
	}
	return interpolate(sorted, (n - 1) * k);
}

/** Exclusive percentile of sorted data: interpolation at rank (n + 1) k - 1. */
export function percentileExc(sorted: readonly number[], k: number): NumResult {
	const n = sorted.length;
	if (n === 0 || k <= 0 || k >= 1) {
		return ERRORS.NUM;
	}
	const h = (n + 1) * k - 1;
	if (h < 0 || h > n - 1) {
		return ERRORS.NUM;
	}
	return interpolate(sorted, h);
}

function interpolate(sorted: readonly number[], h: number): number {
	const i = Math.floor(h);
	const frac = h - i;
	if (i + 1 >= sorted.length || frac === 0) {
		return sorted[i];
	}
	return sorted[i] + frac * (sorted[i + 1] - sorted[i]);
}

/**
 * Relative standing of x in sorted data (LibreOffice / Excel algorithm):
 * an exact match ranks by the count of smaller values, a value between two
 * data points interpolates from the last occurrence of the lower one.
 */
export function percentRank(sorted: readonly number[], x: number, inclusive: boolean): NumResult {
	const n = sorted.length;
	if (n === 0) {
		return ERRORS.NUM;
	}
	if (x < sorted[0] || x > sorted[n - 1]) {
		return ERRORS.NA;
	}
	if (n === 1) {
		return inclusive ? 1 : 0.5;
	}
	let smaller = 0;
	while (smaller < n && sorted[smaller] < x) {
		smaller++;
	}
	if (sorted[smaller] === x) {
		return inclusive ? smaller / (n - 1) : (smaller + 1) / (n + 1);
	}
	// sorted[smaller - 1] < x < sorted[smaller]
	const lower = sorted[smaller - 1];
	const fraction = (x - lower) / (sorted[smaller] - lower);
	return inclusive ? (smaller - 1 + fraction) / (n - 1) : (smaller + fraction) / (n + 1);
}

/** Truncate (not round) to `digits` decimal places, ignoring binary noise. */
function truncateDigits(r: number, digits: number): number {
	const scale = Math.pow(10, digits);
	return approxFloor(r * scale) / scale;
}

type Kernel = (sorted: number[], v: number[]) => NumResult;

function arrayStat(name: string, spec: readonly ArgSpec[], kernel: Kernel): FunctionDefinition {
	const minArgs = 1 + spec.filter((s) => typeof s === 'string').length;
	return define({
		name,
		minArgs,
		maxArgs: 1 + spec.length,
		params: ['any', 'value'],
		impl: (args: FunctionArg[], ctx: FunctionContext) => {
			const xs = numbersOf([args[0]], ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const v = coerceArgs(spec, args, ctx, 1);
			if (isCellError(v)) {
				return v;
			}
			return finish(kernel(sortedAscending(xs), v));
		},
	});
}

const inc: Kernel = (s, [k]) => percentileInc(s, k);
const exc: Kernel = (s, [k]) => percentileExc(s, k);
const quartileInc: Kernel = (s, [q]) => (q < 0 || q > 4 ? ERRORS.NUM : percentileInc(s, q / 4));
const quartileExc: Kernel = (s, [q]) => (q <= 0 || q >= 4 ? ERRORS.NUM : percentileExc(s, q / 4));

function rankKernel(inclusive: boolean): Kernel {
	return (s, [x, digits]) => {
		if (digits < 1) {
			return ERRORS.NUM;
		}
		const r = percentRank(s, x, inclusive);
		return isCellError(r) ? r : truncateDigits(r, digits);
	};
}

const RANK_SPEC: readonly ArgSpec[] = ['n', { i: 3 }];

export const PERCENTILE_FUNCTIONS: FunctionDefinition[] = [
	arrayStat('PERCENTILE', ['n'], inc),
	arrayStat('PERCENTILE.INC', ['n'], inc),
	arrayStat('PERCENTILE.EXC', ['n'], exc),
	arrayStat('QUARTILE', ['i'], quartileInc),
	arrayStat('QUARTILE.INC', ['i'], quartileInc),
	arrayStat('QUARTILE.EXC', ['i'], quartileExc),
	arrayStat('PERCENTRANK', RANK_SPEC, rankKernel(true)),
	arrayStat('PERCENTRANK.INC', RANK_SPEC, rankKernel(true)),
	arrayStat('PERCENTRANK.EXC', RANK_SPEC, rankKernel(false)),
];

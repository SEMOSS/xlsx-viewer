import type { CellError } from '../../../../types/cell';
import { ERRORS, isCellError } from '../../../../values/errors';
import type {
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
} from '../../../types';
import { coerceArgs, define, finish, type NumResult } from '../args';
import { mean } from '../descriptive/core';
import { numericPairs } from '../pairs';

/** Centred sums of paired data (two-pass, so large offsets do not cancel). */
export interface PairSums {
	n: number;
	meanA: number;
	meanB: number;
	saa: number;
	sbb: number;
	sab: number;
}

export function pairSums(as: readonly number[], bs: readonly number[]): PairSums {
	const n = as.length;
	const meanA = mean(as);
	const meanB = mean(bs);
	let saa = 0;
	let sbb = 0;
	let sab = 0;
	for (let i = 0; i < n; i++) {
		const da = as[i] - meanA;
		const db = bs[i] - meanB;
		saa += da * da;
		sbb += db * db;
		sab += da * db;
	}
	return { n, meanA, meanB, saa, sbb, sab };
}

type Kernel = (s: PairSums) => NumResult;

/**
 * A function of two paired arrays. `emptyError` is what Excel documents for
 * no usable pairs (#DIV/0! for the correlation and covariance family, #N/A
 * for the regression family).
 */
function pairFunction(name: string, emptyError: CellError, kernel: Kernel): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['any', 'any'],
		impl: (args, ctx) => evaluatePairs(args[0], args[1], ctx, emptyError, kernel),
	});
}

function evaluatePairs(
	a: FunctionArg,
	b: FunctionArg,
	ctx: FunctionContext,
	emptyError: CellError,
	kernel: Kernel,
): FormulaValue {
	const pairs = numericPairs(a, b, ctx);
	if (isCellError(pairs)) {
		return pairs;
	}
	if (pairs.xs.length === 0) {
		return emptyError;
	}
	return finish(kernel(pairSums(pairs.xs, pairs.ys)));
}

function correlation(s: PairSums): NumResult {
	if (s.saa === 0 || s.sbb === 0) {
		return ERRORS.DIV0;
	}
	return s.sab / Math.sqrt(s.saa * s.sbb);
}

/** Kernels below take (known_y's, known_x's): A is y and B is x. */
function slope(s: PairSums): NumResult {
	return s.sbb === 0 ? ERRORS.DIV0 : s.sab / s.sbb;
}

function intercept(s: PairSums): NumResult {
	return s.sbb === 0 ? ERRORS.DIV0 : s.meanA - (s.sab / s.sbb) * s.meanB;
}

function steyx(s: PairSums): NumResult {
	if (s.n < 3 || s.sbb === 0) {
		return ERRORS.DIV0;
	}
	const resid = Math.max(0, s.saa - (s.sab * s.sab) / s.sbb);
	return Math.sqrt(resid / (s.n - 2));
}

/** FORECAST(x, known_y's, known_x's). */
function forecast(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 3,
		maxArgs: 3,
		params: ['value', 'any', 'any'],
		impl: (args, ctx) => {
			const v = coerceArgs(['n'], args, ctx);
			if (isCellError(v)) {
				return v;
			}
			return evaluatePairs(args[1], args[2], ctx, ERRORS.NA, (s) =>
				s.sbb === 0 ? ERRORS.DIV0 : s.meanA + (s.sab / s.sbb) * (v[0] - s.meanB),
			);
		},
	});
}

export const CORRELATION_FUNCTIONS: FunctionDefinition[] = [
	pairFunction('CORREL', ERRORS.DIV0, correlation),
	pairFunction('PEARSON', ERRORS.NA, correlation),
	pairFunction('RSQ', ERRORS.NA, (s) => {
		const r = correlation(s);
		return isCellError(r) ? r : r * r;
	}),
	pairFunction('COVAR', ERRORS.DIV0, (s) => s.sab / s.n),
	pairFunction('COVARIANCE.P', ERRORS.DIV0, (s) => s.sab / s.n),
	pairFunction('COVARIANCE.S', ERRORS.DIV0, (s) => (s.n < 2 ? ERRORS.DIV0 : s.sab / (s.n - 1))),
	pairFunction('SLOPE', ERRORS.NA, slope),
	pairFunction('INTERCEPT', ERRORS.NA, intercept),
	pairFunction('STEYX', ERRORS.NA, steyx),
	forecast('FORECAST'),
	forecast('FORECAST.LINEAR'),
];

import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { aggregateFunction, type NumResult, scalarFunction } from '../args';
import { A_OPTIONS } from './averages';
import { devsq, mean, standardizedMoment, variance } from './core';

/** Variance (sample or population), #DIV/0! when there are too few points. */
function varianceOf(sample: boolean): (xs: number[]) => NumResult {
	return (xs) => {
		const v = variance(xs, sample);
		return Number.isNaN(v) ? ERRORS.DIV0 : v;
	};
}

function stdevOf(sample: boolean): (xs: number[]) => NumResult {
	return (xs) => {
		const v = variance(xs, sample);
		return Number.isNaN(v) ? ERRORS.DIV0 : Math.sqrt(v);
	};
}

/** Sample excess kurtosis (KURT). */
function kurt(xs: number[]): NumResult {
	const n = xs.length;
	if (n < 4) {
		return ERRORS.DIV0;
	}
	const m = mean(xs);
	const s = Math.sqrt(devsq(xs, m) / (n - 1));
	if (s === 0) {
		return ERRORS.DIV0;
	}
	const m4 = standardizedMoment(xs, m, s, 4);
	return (
		((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * m4 -
		(3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3))
	);
}

/** Sample skewness (SKEW) or population skewness (SKEW.P). */
function skewOf(sample: boolean): (xs: number[]) => NumResult {
	return (xs) => {
		const n = xs.length;
		if (n < 3) {
			return ERRORS.DIV0;
		}
		const m = mean(xs);
		const s = Math.sqrt(devsq(xs, m) / (sample ? n - 1 : n));
		if (s === 0) {
			return ERRORS.DIV0;
		}
		const m3 = standardizedMoment(xs, m, s, 3);
		return sample ? (n / ((n - 1) * (n - 2))) * m3 : m3 / n;
	};
}

export const DISPERSION_FUNCTIONS: FunctionDefinition[] = [
	aggregateFunction('VAR', varianceOf(true)),
	aggregateFunction('VAR.S', varianceOf(true)),
	aggregateFunction('VARA', varianceOf(true), A_OPTIONS),
	aggregateFunction('VARP', varianceOf(false)),
	aggregateFunction('VAR.P', varianceOf(false)),
	aggregateFunction('VARPA', varianceOf(false), A_OPTIONS),
	aggregateFunction('STDEV', stdevOf(true)),
	aggregateFunction('STDEV.S', stdevOf(true)),
	aggregateFunction('STDEVA', stdevOf(true), A_OPTIONS),
	aggregateFunction('STDEVP', stdevOf(false)),
	aggregateFunction('STDEV.P', stdevOf(false)),
	aggregateFunction('STDEVPA', stdevOf(false), A_OPTIONS),
	aggregateFunction('KURT', kurt),
	aggregateFunction('SKEW', skewOf(true)),
	aggregateFunction('SKEW.P', skewOf(false)),
	scalarFunction('STANDARDIZE', ['n', 'n', 'n'], ([x, m, sd]) =>
		sd <= 0 ? ERRORS.NUM : (x - m) / sd,
	),
];

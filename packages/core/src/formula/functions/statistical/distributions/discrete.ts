import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { poissonPmf } from '../special/deviance';
import {
	binomCdf,
	binomPmf,
	binomQuantile,
	binomRange,
	hypergeomCdf,
	hypergeomPmf,
	negBinomCdf,
	negBinomPmf,
	poissonCdf,
} from '../special/discrete';

// Counts are truncated to integers by the `i` argument specs.

function binomDist(k: number, n: number, p: number, cumulative: number): NumResult {
	if (k < 0 || k > n || p < 0 || p > 1) {
		return ERRORS.NUM;
	}
	return cumulative ? binomCdf(k, n, p, 1 - p) : binomPmf(k, n, p, 1 - p);
}

function binomDistRange(n: number, p: number, lo: number, hiArg: number): NumResult {
	const hi = Number.isNaN(hiArg) ? lo : hiArg;
	if (n < 0 || p < 0 || p > 1 || lo < 0 || lo > n || hi < lo || hi > n) {
		return ERRORS.NUM;
	}
	return binomRange(lo, hi, n, p, 1 - p);
}

function binomInv(n: number, p: number, alpha: number): NumResult {
	if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) {
		return ERRORS.NUM;
	}
	return binomQuantile(n, p, alpha);
}

function negBinomDist(f: number, s: number, p: number, cumulative: number): NumResult {
	if (p < 0 || p > 1 || f < 0 || s < 1) {
		return ERRORS.NUM;
	}
	return cumulative ? negBinomCdf(f, s, p, 1 - p) : negBinomPmf(f, s, p, 1 - p);
}

/** HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative). */
function hypgeomDist(
	x: number,
	n: number,
	m: number,
	total: number,
	cumulative: number,
): NumResult {
	if (
		x < 0 ||
		x > Math.min(n, m) ||
		x < Math.max(0, n - total + m) ||
		n <= 0 ||
		n > total ||
		m <= 0 ||
		m > total ||
		total <= 0
	) {
		return ERRORS.NUM;
	}
	return cumulative ? hypergeomCdf(x, n, m, total) : hypergeomPmf(x, n, m, total);
}

function poissonDist(k: number, mean: number, cumulative: number): NumResult {
	if (k < 0 || mean < 0) {
		return ERRORS.NUM;
	}
	return cumulative ? poissonCdf(k, mean) : poissonPmf(k, mean);
}

/** PERMUT(n, k) = n! / (n - k)!. */
function permut(n: number, k: number): NumResult {
	if (n <= 0 || k < 0 || n < k) {
		return ERRORS.NUM;
	}
	let r = 1;
	for (let i = 0; i < k && Number.isFinite(r); i++) {
		r *= n - i;
	}
	return r;
}

function permutationA(n: number, k: number): NumResult {
	return n < 0 || k < 0 ? ERRORS.NUM : Math.pow(n, k);
}

export const DISCRETE_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('BINOM.DIST', ['i', 'i', 'n', 'b'], ([k, n, p, c]) => binomDist(k, n, p, c)),
	scalarFunction('BINOMDIST', ['i', 'i', 'n', 'b'], ([k, n, p, c]) => binomDist(k, n, p, c)),
	scalarFunction('BINOM.DIST.RANGE', ['i', 'n', 'i', { i: NaN }], ([n, p, lo, hi]) =>
		binomDistRange(n, p, lo, hi),
	),
	scalarFunction('BINOM.INV', ['i', 'n', 'n'], ([n, p, a]) => binomInv(n, p, a)),
	scalarFunction('CRITBINOM', ['i', 'n', 'n'], ([n, p, a]) => binomInv(n, p, a)),
	scalarFunction('NEGBINOM.DIST', ['i', 'i', 'n', 'b'], ([f, s, p, c]) => negBinomDist(f, s, p, c)),
	scalarFunction('NEGBINOMDIST', ['i', 'i', 'n'], ([f, s, p]) => negBinomDist(f, s, p, 0)),
	scalarFunction('HYPGEOM.DIST', ['i', 'i', 'i', 'i', 'b'], ([x, n, m, t, c]) =>
		hypgeomDist(x, n, m, t, c),
	),
	scalarFunction('HYPGEOMDIST', ['i', 'i', 'i', 'i'], ([x, n, m, t]) => hypgeomDist(x, n, m, t, 0)),
	scalarFunction('POISSON.DIST', ['i', 'n', 'b'], ([k, m, c]) => poissonDist(k, m, c)),
	scalarFunction('POISSON', ['i', 'n', 'b'], ([k, m, c]) => poissonDist(k, m, c)),
	scalarFunction('PERMUT', ['i', 'i'], ([n, k]) => permut(n, k)),
	scalarFunction('PERMUTATIONA', ['i', 'i'], ([n, k]) => permutationA(n, k)),
];

import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { gamma, lnGamma } from '../special/gamma-function';
import { gammaDensity, gammaQuantile, incGamma } from '../special/incomplete-gamma';

function gammaDist(x: number, alpha: number, beta: number, cumulative: number): NumResult {
	if (x < 0 || alpha <= 0 || beta <= 0) {
		return ERRORS.NUM;
	}
	const z = x / beta;
	return cumulative ? incGamma(alpha, z).lower : gammaDensity(alpha, z) / beta;
}

function gammaInv(p: number, alpha: number, beta: number): NumResult {
	if (p < 0 || p > 1 || alpha <= 0 || beta <= 0) {
		return ERRORS.NUM;
	}
	return gammaQuantile(alpha, p) * beta;
}

function exponDist(x: number, lambda: number, cumulative: number): NumResult {
	if (x < 0 || lambda <= 0) {
		return ERRORS.NUM;
	}
	return cumulative ? -Math.expm1(-lambda * x) : lambda * Math.exp(-lambda * x);
}

function weibullDist(x: number, alpha: number, beta: number, cumulative: number): NumResult {
	if (x < 0 || alpha <= 0 || beta <= 0) {
		return ERRORS.NUM;
	}
	const t = Math.pow(x / beta, alpha);
	if (cumulative) {
		return -Math.expm1(-t);
	}
	return (alpha / beta) * Math.pow(x / beta, alpha - 1) * Math.exp(-t);
}

function lnGammaOf(x: number): NumResult {
	return x <= 0 ? ERRORS.NUM : lnGamma(x);
}

export const GAMMA_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('GAMMA', ['n'], ([x]) => gamma(x)),
	scalarFunction('GAMMALN', ['n'], ([x]) => lnGammaOf(x)),
	scalarFunction('GAMMALN.PRECISE', ['n'], ([x]) => lnGammaOf(x)),
	scalarFunction('GAMMA.DIST', ['n', 'n', 'n', 'b'], ([x, a, b, c]) => gammaDist(x, a, b, c)),
	scalarFunction('GAMMADIST', ['n', 'n', 'n', 'b'], ([x, a, b, c]) => gammaDist(x, a, b, c)),
	scalarFunction('GAMMA.INV', ['n', 'n', 'n'], ([p, a, b]) => gammaInv(p, a, b)),
	scalarFunction('GAMMAINV', ['n', 'n', 'n'], ([p, a, b]) => gammaInv(p, a, b)),
	scalarFunction('EXPON.DIST', ['n', 'n', 'b'], ([x, l, c]) => exponDist(x, l, c)),
	scalarFunction('EXPONDIST', ['n', 'n', 'b'], ([x, l, c]) => exponDist(x, l, c)),
	scalarFunction('WEIBULL', ['n', 'n', 'n', 'b'], ([x, a, b, c]) => weibullDist(x, a, b, c)),
	scalarFunction('WEIBULL.DIST', ['n', 'n', 'n', 'b'], ([x, a, b, c]) => weibullDist(x, a, b, c)),
];

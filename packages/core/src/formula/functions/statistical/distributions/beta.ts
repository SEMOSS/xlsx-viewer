import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { betaQuantile } from '../special/beta-quantile';
import { betaDensity, incBeta } from '../special/incomplete-beta';

function betaDist(
	x: number,
	alpha: number,
	beta: number,
	cumulative: number,
	lo: number,
	hi: number,
): NumResult {
	if (alpha <= 0 || beta <= 0 || x < lo || x > hi || lo === hi) {
		return ERRORS.NUM;
	}
	const width = hi - lo;
	const z = (x - lo) / width;
	const y = (hi - x) / width;
	return cumulative ? incBeta(alpha, beta, z, y).lower : betaDensity(alpha, beta, z, y) / width;
}

function betaInv(p: number, alpha: number, beta: number, lo: number, hi: number): NumResult {
	if (p <= 0 || p > 1 || alpha <= 0 || beta <= 0 || lo >= hi) {
		return ERRORS.NUM;
	}
	const { x, y } = betaQuantile(alpha, beta, p);
	const width = hi - lo;
	// Scale from the end the quantile is closer to, keeping its precision.
	return x <= 0.5 ? lo + x * width : hi - y * width;
}

export const BETA_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('BETA.DIST', ['n', 'n', 'n', 'b', { n: 0 }, { n: 1 }], ([x, a, b, c, lo, hi]) =>
		betaDist(x, a, b, c, lo, hi),
	),
	scalarFunction('BETADIST', ['n', 'n', 'n', { n: 0 }, { n: 1 }], ([x, a, b, lo, hi]) =>
		betaDist(x, a, b, 1, lo, hi),
	),
	scalarFunction('BETA.INV', ['n', 'n', 'n', { n: 0 }, { n: 1 }], ([p, a, b, lo, hi]) =>
		betaInv(p, a, b, lo, hi),
	),
	scalarFunction('BETAINV', ['n', 'n', 'n', { n: 0 }, { n: 1 }], ([p, a, b, lo, hi]) =>
		betaInv(p, a, b, lo, hi),
	),
];

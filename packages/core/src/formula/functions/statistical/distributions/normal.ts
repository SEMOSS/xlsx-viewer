import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { erf } from '../special/erf';
import { normCdf, normInv, normPdf } from '../special/normal';

function normDist(x: number, mean: number, sd: number, cumulative: number): NumResult {
	if (sd <= 0) {
		return ERRORS.NUM;
	}
	const z = (x - mean) / sd;
	return cumulative ? normCdf(z) : normPdf(z) / sd;
}

function normQuantile(p: number, mean: number, sd: number): NumResult {
	if (p <= 0 || p >= 1 || sd <= 0) {
		return ERRORS.NUM;
	}
	return mean + sd * normInv(p);
}

function standardQuantile(p: number): NumResult {
	return p <= 0 || p >= 1 ? ERRORS.NUM : normInv(p);
}

function lognormDist(x: number, mean: number, sd: number, cumulative: number): NumResult {
	if (x <= 0 || sd <= 0) {
		return ERRORS.NUM;
	}
	const z = (Math.log(x) - mean) / sd;
	return cumulative ? normCdf(z) : normPdf(z) / (x * sd);
}

function lognormQuantile(p: number, mean: number, sd: number): NumResult {
	if (p <= 0 || p >= 1 || sd <= 0) {
		return ERRORS.NUM;
	}
	return Math.exp(mean + sd * normInv(p));
}

/** Half-width of the normal confidence interval: z(1 - alpha / 2) sd / sqrt(n). */
export function confidenceNorm(alpha: number, sd: number, size: number): NumResult {
	if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) {
		return ERRORS.NUM;
	}
	return (-normInv(alpha / 2) * sd) / Math.sqrt(size);
}

export const NORMAL_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('NORM.DIST', ['n', 'n', 'n', 'b'], ([x, m, s, c]) => normDist(x, m, s, c)),
	scalarFunction('NORMDIST', ['n', 'n', 'n', 'b'], ([x, m, s, c]) => normDist(x, m, s, c)),
	scalarFunction('NORM.INV', ['n', 'n', 'n'], ([p, m, s]) => normQuantile(p, m, s)),
	scalarFunction('NORMINV', ['n', 'n', 'n'], ([p, m, s]) => normQuantile(p, m, s)),
	scalarFunction('NORM.S.DIST', ['n', 'b'], ([z, c]) => (c ? normCdf(z) : normPdf(z))),
	scalarFunction('NORMSDIST', ['n'], ([z]) => normCdf(z)),
	scalarFunction('NORM.S.INV', ['n'], ([p]) => standardQuantile(p)),
	scalarFunction('NORMSINV', ['n'], ([p]) => standardQuantile(p)),
	scalarFunction('PHI', ['n'], ([z]) => normPdf(z)),
	scalarFunction('GAUSS', ['n'], ([z]) => 0.5 * erf(z * Math.SQRT1_2)),
	scalarFunction('LOGNORM.DIST', ['n', 'n', 'n', 'b'], ([x, m, s, c]) => lognormDist(x, m, s, c)),
	scalarFunction('LOGNORMDIST', ['n', 'n', 'n'], ([x, m, s]) => lognormDist(x, m, s, 1)),
	scalarFunction('LOGNORM.INV', ['n', 'n', 'n'], ([p, m, s]) => lognormQuantile(p, m, s)),
	scalarFunction('LOGINV', ['n', 'n', 'n'], ([p, m, s]) => lognormQuantile(p, m, s)),
	scalarFunction('CONFIDENCE', ['n', 'n', 'i'], ([a, s, n]) => confidenceNorm(a, s, n)),
	scalarFunction('CONFIDENCE.NORM', ['n', 'n', 'i'], ([a, s, n]) => confidenceNorm(a, s, n)),
	scalarFunction('FISHER', ['n'], ([x]) => (x <= -1 || x >= 1 ? ERRORS.NUM : Math.atanh(x))),
	scalarFunction('FISHERINV', ['n'], ([y]) => Math.tanh(y)),
];

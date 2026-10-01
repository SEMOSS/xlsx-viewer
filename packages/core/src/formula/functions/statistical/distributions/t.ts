import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { tCdf, tPdf, tQuantile, tQuantileTwoTailed, tSf, tTwoTailed } from '../special/student-t';

// Degrees of freedom are truncated by the `i` specs; Excel requires at least 1.

function tDist(x: number, df: number, cumulative: number): NumResult {
	if (df < 1) {
		return ERRORS.NUM;
	}
	return cumulative ? tCdf(x, df) : tPdf(x, df);
}

function tDist2T(x: number, df: number): NumResult {
	return x < 0 || df < 1 ? ERRORS.NUM : tTwoTailed(x, df);
}

function tDistRT(x: number, df: number): NumResult {
	return df < 1 ? ERRORS.NUM : tSf(x, df);
}

/** Legacy TDIST(x, df, tails). */
function tdist(x: number, df: number, tails: number): NumResult {
	if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) {
		return ERRORS.NUM;
	}
	return tails === 1 ? tSf(x, df) : tTwoTailed(x, df);
}

function tInv(p: number, df: number): NumResult {
	if (p <= 0 || p >= 1 || df < 1) {
		return ERRORS.NUM;
	}
	return tQuantile(p, df);
}

function tInv2T(p: number, df: number): NumResult {
	if (p <= 0 || p > 1 || df < 1) {
		return ERRORS.NUM;
	}
	return tQuantileTwoTailed(p, df);
}

/** Half-width of the Student t confidence interval for the mean. */
function confidenceT(alpha: number, sd: number, size: number): NumResult {
	if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) {
		return ERRORS.NUM;
	}
	if (size === 1) {
		return ERRORS.DIV0;
	}
	return (tQuantileTwoTailed(alpha, size - 1) * sd) / Math.sqrt(size);
}

export const T_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('T.DIST', ['n', 'i', 'b'], ([x, df, c]) => tDist(x, df, c)),
	scalarFunction('T.DIST.2T', ['n', 'i'], ([x, df]) => tDist2T(x, df)),
	scalarFunction('T.DIST.RT', ['n', 'i'], ([x, df]) => tDistRT(x, df)),
	scalarFunction('TDIST', ['n', 'i', 'i'], ([x, df, tails]) => tdist(x, df, tails)),
	scalarFunction('T.INV', ['n', 'i'], ([p, df]) => tInv(p, df)),
	scalarFunction('T.INV.2T', ['n', 'i'], ([p, df]) => tInv2T(p, df)),
	scalarFunction('TINV', ['n', 'i'], ([p, df]) => tInv2T(p, df)),
	scalarFunction('CONFIDENCE.T', ['n', 'n', 'i'], ([a, s, n]) => confidenceT(a, s, n)),
];

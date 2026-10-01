import { ERRORS } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { type NumResult, scalarFunction } from '../args';
import { fPdf, fQuantile, fTails } from '../special/f-distribution';

/**
 * Degrees of freedom are truncated and must be at least 1; the right-tail
 * and legacy forms also reject 10^10 or more, as documented by Excel.
 */
function badDf(d1: number, d2: number, capped: boolean): boolean {
	if (d1 < 1 || d2 < 1) {
		return true;
	}
	return capped && (d1 >= 1e10 || d2 >= 1e10);
}

function fDist(x: number, d1: number, d2: number, cumulative: number): NumResult {
	if (x < 0 || badDf(d1, d2, false)) {
		return ERRORS.NUM;
	}
	return cumulative ? fTails(x, d1, d2).lower : fPdf(x, d1, d2);
}

function fRight(x: number, d1: number, d2: number): NumResult {
	if (x < 0 || badDf(d1, d2, true)) {
		return ERRORS.NUM;
	}
	return fTails(x, d1, d2).upper;
}

function fInv(p: number, d1: number, d2: number, upperTail: boolean): NumResult {
	if (p < 0 || p > 1 || badDf(d1, d2, upperTail)) {
		return ERRORS.NUM;
	}
	return fQuantile(p, d1, d2, upperTail);
}

export const F_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('F.DIST', ['n', 'i', 'i', 'b'], ([x, a, b, c]) => fDist(x, a, b, c)),
	scalarFunction('F.DIST.RT', ['n', 'i', 'i'], ([x, a, b]) => fRight(x, a, b)),
	scalarFunction('FDIST', ['n', 'i', 'i'], ([x, a, b]) => fRight(x, a, b)),
	scalarFunction('F.INV', ['n', 'i', 'i'], ([p, a, b]) => fInv(p, a, b, false)),
	scalarFunction('F.INV.RT', ['n', 'i', 'i'], ([p, a, b]) => fInv(p, a, b, true)),
	scalarFunction('FINV', ['n', 'i', 'i'], ([p, a, b]) => fInv(p, a, b, true)),
];

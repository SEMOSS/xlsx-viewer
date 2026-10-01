import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { define, finish, type NumResult, scalarFunction } from '../args';
import { gridAt, sameShape, toGrid } from '../counting/grid';
import { gammaDensity, gammaQuantile, incGamma } from '../special/incomplete-gamma';

/** Excel accepts 1 <= df <= 10^10 (truncated) for the chi-squared functions. */
function badDf(df: number): boolean {
	return df < 1 || df > 1e10;
}

function chisqDist(x: number, df: number, cumulative: number): NumResult {
	if (x < 0 || badDf(df)) {
		return ERRORS.NUM;
	}
	const k = df / 2;
	return cumulative ? incGamma(k, x / 2).lower : gammaDensity(k, x / 2) / 2;
}

function chisqRight(x: number, df: number): NumResult {
	if (x < 0 || badDf(df)) {
		return ERRORS.NUM;
	}
	return incGamma(df / 2, x / 2).upper;
}

function chisqInv(p: number, df: number, upperTail: boolean): NumResult {
	if (p < 0 || p > 1 || badDf(df)) {
		return ERRORS.NUM;
	}
	return 2 * gammaQuantile(df / 2, p, upperTail);
}

/** Right-tail probability of Pearson's statistic sum (A - E)^2 / E. */
function chisqTest(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['any', 'any'],
		impl: (args, ctx) => {
			const actual = toGrid(args[0], ctx);
			if (isCellError(actual)) {
				return actual;
			}
			const expected = toGrid(args[1], ctx);
			if (isCellError(expected)) {
				return expected;
			}
			if (!sameShape(actual, expected)) {
				return ERRORS.NA;
			}
			const { rows, cols } = actual;
			const df = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : rows * cols - 1;
			if (df < 1) {
				return ERRORS.NA;
			}
			let chi = 0;
			const lastRow = Math.min(rows, Math.max(actual.filledRows, expected.filledRows));
			for (let r = 0; r < lastRow; r++) {
				for (let c = 0; c < cols; c++) {
					const a = gridAt(actual, r, c);
					const e = gridAt(expected, r, c);
					if (isCellError(a)) {
						return a;
					}
					if (isCellError(e)) {
						return e;
					}
					if (typeof a !== 'number' || typeof e !== 'number') {
						continue;
					}
					if (e === 0) {
						return ERRORS.DIV0;
					}
					chi += ((a - e) * (a - e)) / e;
				}
			}
			if (chi < 0) {
				return ERRORS.NUM;
			}
			return finish(incGamma(df / 2, chi / 2).upper);
		},
	});
}

export const CHISQ_FUNCTIONS: FunctionDefinition[] = [
	scalarFunction('CHISQ.DIST', ['n', 'i', 'b'], ([x, df, c]) => chisqDist(x, df, c)),
	scalarFunction('CHISQ.DIST.RT', ['n', 'i'], ([x, df]) => chisqRight(x, df)),
	scalarFunction('CHIDIST', ['n', 'i'], ([x, df]) => chisqRight(x, df)),
	scalarFunction('CHISQ.INV', ['n', 'i'], ([p, df]) => chisqInv(p, df, false)),
	scalarFunction('CHISQ.INV.RT', ['n', 'i'], ([p, df]) => chisqInv(p, df, true)),
	scalarFunction('CHIINV', ['n', 'i'], ([p, df]) => chisqInv(p, df, true)),
	chisqTest('CHISQ.TEST'),
	chisqTest('CHITEST'),
];

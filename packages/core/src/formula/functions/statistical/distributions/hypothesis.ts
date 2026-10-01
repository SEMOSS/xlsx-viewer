import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { arrayNumbers, coerceArgs, define, finish, isMissing, type NumResult } from '../args';
import { devsq, mean } from '../descriptive/core';
import { numericPairs } from '../pairs';
import { fTails } from '../special/f-distribution';
import { normSf } from '../special/normal';
import { tSf, tTwoTailed } from '../special/student-t';

/** p-value of a t statistic for one or two tails (real degrees of freedom allowed). */
function tPValue(t: number, df: number, tails: number): number {
	return tails === 1 ? tSf(Math.abs(t), df) : tTwoTailed(t, df);
}

function pairedT(diffs: number[]): { t: number; df: number } | null {
	const n = diffs.length;
	if (n < 2) {
		return null;
	}
	const m = mean(diffs);
	const se = Math.sqrt(devsq(diffs, m) / (n - 1) / n);
	return se === 0 ? null : { t: m / se, df: n - 1 };
}

function twoSampleT(
	xs: number[],
	ys: number[],
	equalVariance: boolean,
): { t: number; df: number } | null {
	const n1 = xs.length;
	const n2 = ys.length;
	if (n1 < 2 || n2 < 2) {
		return null;
	}
	const m1 = mean(xs);
	const m2 = mean(ys);
	const ss1 = devsq(xs, m1);
	const ss2 = devsq(ys, m2);
	if (equalVariance) {
		const pooled = (ss1 + ss2) / (n1 + n2 - 2);
		const se = Math.sqrt(pooled * (1 / n1 + 1 / n2));
		return se === 0 ? null : { t: (m1 - m2) / se, df: n1 + n2 - 2 };
	}
	// Welch's test with the Satterthwaite degrees of freedom (not truncated).
	const v1 = ss1 / (n1 - 1) / n1;
	const v2 = ss2 / (n2 - 1) / n2;
	const v = v1 + v2;
	if (v === 0) {
		return null;
	}
	const df = (v * v) / ((v1 * v1) / (n1 - 1) + (v2 * v2) / (n2 - 1));
	return { t: (m1 - m2) / Math.sqrt(v), df };
}

function tTest(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 4,
		maxArgs: 4,
		params: ['any', 'any', 'value', 'value'],
		impl: (args, ctx) => {
			const v = coerceArgs(['i', 'i'], args, ctx, 2);
			if (isCellError(v)) {
				return v;
			}
			const [tails, type] = v;
			if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) {
				return ERRORS.NUM;
			}
			let stat: { t: number; df: number } | null;
			if (type === 1) {
				const pairs = numericPairs(args[0], args[1], ctx);
				if (isCellError(pairs)) {
					return pairs;
				}
				stat = pairedT(pairs.xs.map((x, i) => x - pairs.ys[i]));
			} else {
				const xs = arrayNumbers(args[0], ctx);
				if (isCellError(xs)) {
					return xs;
				}
				const ys = arrayNumbers(args[1], ctx);
				if (isCellError(ys)) {
					return ys;
				}
				stat = twoSampleT(xs, ys, type === 2);
			}
			return stat ? finish(tPValue(stat.t, stat.df, tails)) : ERRORS.DIV0;
		},
	});
}

function fTest(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['any', 'any'],
		impl: (args, ctx) => {
			const xs = arrayNumbers(args[0], ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const ys = arrayNumbers(args[1], ctx);
			if (isCellError(ys)) {
				return ys;
			}
			if (xs.length < 2 || ys.length < 2) {
				return ERRORS.DIV0;
			}
			const v1 = devsq(xs) / (xs.length - 1);
			const v2 = devsq(ys) / (ys.length - 1);
			if (v1 === 0 || v2 === 0) {
				return ERRORS.DIV0;
			}
			const tails = fTails(v1 / v2, xs.length - 1, ys.length - 1);
			return finish(2 * Math.min(tails.lower, tails.upper));
		},
	});
}

/** One-tailed p-value of the z statistic (mean - x) / (sigma / sqrt(n)). */
function zTest(name: string): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 3,
		params: ['any', 'value', 'value'],
		impl: (args, ctx) => {
			const xs = arrayNumbers(args[0], ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const v = coerceArgs(['n', { n: NaN }], args, ctx, 1);
			if (isCellError(v)) {
				return v;
			}
			const n = xs.length;
			if (n === 0) {
				return ERRORS.NA;
			}
			const m = mean(xs);
			let sigma: NumResult;
			if (isMissing(args[2]) || Number.isNaN(v[1])) {
				sigma = n < 2 ? ERRORS.DIV0 : Math.sqrt(devsq(xs, m) / (n - 1));
			} else {
				sigma = v[1] <= 0 ? ERRORS.NUM : v[1];
			}
			if (isCellError(sigma)) {
				return sigma;
			}
			if (sigma === 0) {
				return ERRORS.DIV0;
			}
			return finish(normSf((m - v[0]) / (sigma / Math.sqrt(n))));
		},
	});
}

export const HYPOTHESIS_FUNCTIONS: FunctionDefinition[] = [
	tTest('T.TEST'),
	tTest('TTEST'),
	fTest('F.TEST'),
	fTest('FTEST'),
	zTest('Z.TEST'),
	zTest('ZTEST'),
];

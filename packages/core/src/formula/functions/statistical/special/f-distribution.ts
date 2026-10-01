import { betaQuantile } from './beta-quantile';
import type { TailPair } from './incomplete-gamma';
import { betaPrefix, incBeta } from './incomplete-beta';

/**
 * Fisher's F distribution with d1 and d2 degrees of freedom, through the
 * Beta(d1 / 2, d2 / 2) variable x = d1 F / (d1 F + d2).
 */

/** P(F <= f) and P(F > f). */
export function fTails(f: number, d1: number, d2: number): TailPair {
	if (Number.isNaN(f)) {
		return { lower: NaN, upper: NaN };
	}
	if (f <= 0) {
		return { lower: 0, upper: 1 };
	}
	if (f === Infinity) {
		return { lower: 1, upper: 0 };
	}
	const denom = d1 * f + d2;
	return incBeta(d1 / 2, d2 / 2, (d1 * f) / denom, d2 / denom);
}

/** Density of the F distribution. */
export function fPdf(f: number, d1: number, d2: number): number {
	if (f < 0) {
		return 0;
	}
	if (f === 0) {
		return d1 < 2 ? Infinity : d1 === 2 ? 1 : 0;
	}
	const denom = d1 * f + d2;
	return betaPrefix(d1 / 2, d2 / 2, (d1 * f) / denom, d2 / denom) / f;
}

/** Quantile at `prob` (lower tail, or upper tail when `upperTail` is set). */
export function fQuantile(prob: number, d1: number, d2: number, upperTail = false): number {
	const { x, y } = betaQuantile(d1 / 2, d2 / 2, prob, upperTail);
	if (x === 0) {
		return 0;
	}
	return (d2 * x) / (d1 * y);
}

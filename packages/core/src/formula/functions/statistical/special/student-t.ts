import { betaQuantile } from './beta-quantile';
import { incBeta, lnBeta } from './incomplete-beta';

/**
 * Student's t distribution for real degrees of freedom df > 0, through the
 * incomplete beta function: P(|T| > t) = I_x(df / 2, 1 / 2), x = df / (df + t^2).
 */

/** Two-tailed probability P(|T| > |t|). */
export function tTwoTailed(t: number, df: number): number {
	if (Number.isNaN(t)) {
		return NaN;
	}
	const at = Math.abs(t);
	if (at === 0) {
		return 1;
	}
	if (df === 1) {
		return (2 * Math.atan2(1, at)) / Math.PI;
	}
	const t2 = at * at;
	if (!Number.isFinite(t2)) {
		return 0;
	}
	const denom = df + t2;
	return incBeta(df / 2, 0.5, df / denom, t2 / denom).lower;
}

/** Upper tail P(T > t). */
export function tSf(t: number, df: number): number {
	if (t < 0) {
		return 1 - tTwoTailed(t, df) / 2;
	}
	return tTwoTailed(t, df) / 2;
}

/** Distribution function P(T <= t). */
export function tCdf(t: number, df: number): number {
	return tSf(-t, df);
}

/** Density of Student's t distribution. */
export function tPdf(t: number, df: number): number {
	const z = (t * t) / df;
	if (!Number.isFinite(z)) {
		return 0;
	}
	return Math.exp(-((df + 1) / 2) * Math.log1p(z) - lnBeta(df / 2, 0.5)) / Math.sqrt(df);
}

/** t >= 0 with P(|T| > t) = p, for 0 < p <= 1. */
export function tQuantileTwoTailed(p: number, df: number): number {
	if (p >= 1) {
		return 0;
	}
	if (df === 1) {
		// t = cot(pi p / 2); use tan of the complement when p is large.
		return p > 0.5 ? Math.tan((Math.PI * (1 - p)) / 2) : 1 / Math.tan((Math.PI * p) / 2);
	}
	if (df === 2) {
		return (Math.SQRT2 * (1 - p)) / Math.sqrt(p * (2 - p));
	}
	const { x, y } = betaQuantile(df / 2, 0.5, p);
	return Math.sqrt((df * y) / x);
}

/** t with P(T <= t) = p, for 0 < p < 1. */
export function tQuantile(p: number, df: number): number {
	if (p === 0.5) {
		return 0;
	}
	return p < 0.5 ? -tQuantileTwoTailed(2 * p, df) : tQuantileTwoTailed(2 * (1 - p), df);
}

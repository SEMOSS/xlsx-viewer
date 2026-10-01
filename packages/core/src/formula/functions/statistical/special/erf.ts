/**
 * Error function family, accurate to a few ulps over the whole double range.
 *
 * - |x| < 1.5: Kummer's positive-term series
 *   erf(x) = 2x / sqrt(pi) * e^(-x^2) * sum (2x^2)^n / (1 * 3 * ... * (2n + 1)).
 * - erfc for x >= 0.75 (and erf from 1.5): the even continued fraction for
 *   erfc, evaluated bottom-up and scaled by e^(-x^2).
 *
 * e^(-x^2) is evaluated with x^2 split into an exact high part and a small
 * correction, so the tails keep full relative precision.
 */

export const SQRT_PI = 1.772453850905516;
export const SQRT_2PI = 2.5066282746310007;
export const LN_SQRT_2PI = 0.9189385332046728;
const TWO_OVER_SQRT_PI = 1.1283791670955126;
/** erf uses the series below this point. */
const SERIES_CUTOFF = 1.5;
/** erfc uses the continued fraction from this point (1 - erf loses digits beyond it). */
const FRACTION_CUTOFF = 0.75;
/** 2^27 + 1, the Veltkamp splitting constant for doubles. */
const SPLITTER = 134217729;

/** High part of `x` with at most 26 significant bits, so `hi * hi` is exact. */
function highPart(x: number): number {
	const c = SPLITTER * x;
	return c - (c - x);
}

/** e^(-x^2) without the rounding error of forming x^2. */
export function expNegSquare(x: number): number {
	const ax = Math.abs(x);
	if (ax > 27.3) {
		return 0;
	}
	const hi = highPart(ax);
	const lo = ax - hi;
	return Math.exp(-hi * hi) * Math.exp(-lo * (ax + hi));
}

/** e^(-x^2 / 2) without the rounding error of forming x^2. */
export function expNegHalfSquare(x: number): number {
	const ax = Math.abs(x);
	if (ax > 38.6) {
		return 0;
	}
	const hi = highPart(ax);
	const lo = ax - hi;
	return Math.exp((-hi * hi) / 2) * Math.exp((-lo * (ax + hi)) / 2);
}

/** erf(x) for 0 <= x < 1.5 by the positive-term series. */
function erfSeries(x: number): number {
	const z = 2 * x * x;
	let term = 1;
	let sum = 1;
	for (let n = 1; n < 200; n++) {
		term *= z / (2 * n + 1);
		sum += term;
		if (term < 1e-17 * sum) {
			break;
		}
	}
	return TWO_OVER_SQRT_PI * x * expNegSquare(x) * sum;
}

/**
 * e^(x^2) * erfc(x) for x >= 0.75 from the continued fraction
 * erfc(x) = x e^(-x^2) / sqrt(pi) / (b0 - a1 / (b1 - a2 / (b2 - ...))),
 * a_n = n (2n - 1) / 2, b_n = x^2 + 2n + 1/2. It is evaluated bottom-up with
 * a term count chosen for convergence below 1e-17, which keeps the rounding
 * error near one ulp (forward Lentz evaluation accumulates far more).
 */
function erfcScaledFraction(x: number): number {
	if (x > 1e8) {
		return 1 / (x * SQRT_PI);
	}
	const x2 = x * x;
	const terms = Math.ceil(8 + 150 / x2);
	let t = 0;
	for (let n = terms; n >= 1; n--) {
		t = (n * (2 * n - 1)) / 2 / (x2 + 2 * n + 0.5 - t);
	}
	return x / (SQRT_PI * (x2 + 0.5 - t));
}

/** The error function. */
export function erf(x: number): number {
	if (Number.isNaN(x)) {
		return NaN;
	}
	const ax = Math.abs(x);
	let r: number;
	if (ax < SERIES_CUTOFF) {
		r = erfSeries(ax);
	} else if (ax >= 6) {
		r = 1;
	} else {
		r = 1 - erfcScaledFraction(ax) * expNegSquare(ax);
	}
	return x < 0 ? -r : r;
}

/** The complementary error function 1 - erf(x), accurate in the upper tail. */
export function erfc(x: number): number {
	if (Number.isNaN(x)) {
		return NaN;
	}
	if (x < 0) {
		return 2 - erfc(-x);
	}
	if (x < FRACTION_CUTOFF) {
		return 1 - erfSeries(x);
	}
	return erfcScaledFraction(x) * expNegSquare(x);
}

/** The scaled complementary error function e^(x^2) * erfc(x), for x >= 0. */
export function erfcx(x: number): number {
	if (x < FRACTION_CUTOFF) {
		return Math.exp(x * x) * (1 - erfSeries(x));
	}
	return erfcScaledFraction(x);
}

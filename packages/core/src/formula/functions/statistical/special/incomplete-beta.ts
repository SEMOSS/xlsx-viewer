import { log1pmx } from './deviance';
import { LN_SQRT_2PI } from './erf';
import { gamma, lnGamma, lnGammaRatio, stirlerr } from './gamma-function';
import type { TailPair } from './incomplete-gamma';

/**
 * Regularized incomplete beta function I_x(a, b) by a continued fraction
 * applied on whichever side of the mean converges quickly. Callers pass both
 * x and y = 1 - x so that whichever is small keeps its full relative
 * precision.
 */

/** Relative change at which the continued fraction counts as converged (about 2 ulps). */
const EPS = 4e-16;

/** Rounding error of s = a + b, so that a + b = s + e exactly (Knuth's TwoSum). */
function sumError(a: number, b: number, s: number): number {
	const bv = s - a;
	return a - (s - bv) + (b - bv);
}

/** Digamma to about 1e-6 relative: enough for first-order rounding corrections. */
function digammaApprox(x: number): number {
	let r = 0;
	let z = x;
	while (z < 6) {
		r -= 1 / z;
		z += 1;
	}
	return r + Math.log(z) - 1 / (2 * z) - 1 / (12 * z * z);
}

/** ln Gamma(a + b), corrected for the rounding of the sum. */
function lnGammaOfSum(a: number, b: number): number {
	const s = a + b;
	return lnGamma(s) + digammaApprox(s) * sumError(a, b, s);
}

/** ln B(a, b) for a, b > 0. */
export function lnBeta(a: number, b: number): number {
	const lo = Math.min(a, b);
	const hi = Math.max(a, b);
	if (hi < 10) {
		return lnGamma(a) + lnGamma(b) - lnGammaOfSum(a, b);
	}
	if (lo < 10) {
		return lnGamma(lo) - lnGammaRatio(hi, lo);
	}
	const s = a + b;
	return (
		LN_SQRT_2PI +
		(lo - 0.5) * Math.log(lo / s) +
		(hi - 0.5) * Math.log1p(-lo / s) -
		0.5 * Math.log(s) +
		stirlerr(lo) +
		stirlerr(hi) -
		stirlerr(s)
	);
}

/**
 * x^a y^b / B(a, b) with y = 1 - x. For large a and b the deviance form
 * a log1pmx(u) + b log1pmx(v) is used, which avoids subtracting huge
 * logarithms (the binomial saddle point).
 */
export function betaPrefix(a: number, b: number, x: number, y: number): number {
	if (x <= 0 || y <= 0) {
		return 0;
	}
	if (a >= 10 && b >= 10) {
		const s = a + b;
		const p = a / s;
		const q = b / s;
		let u: number;
		let v: number;
		if (x <= y) {
			const d = x - p;
			u = d / p;
			v = -d / q;
		} else {
			const d = y - q;
			v = d / q;
			u = -d / p;
		}
		// Far below the mean 1 + u = x / p is tiny and must come from x itself.
		const lu = u < -0.5 ? Math.log(x / p) - u : log1pmx(u);
		const lv = v < -0.5 ? Math.log(y / q) - v : log1pmx(v);
		const deviance = a * lu + b * lv;
		const correction = stirlerr(s) - stirlerr(a) - stirlerr(b);
		return Math.sqrt((p * b) / (2 * Math.PI)) * Math.exp(deviance + correction);
	}
	if (a + b < 170) {
		// Powers of the smaller variable directly, of the larger one through log1p of the
		// smaller (exact for the input), and B(a, b) from gamma values: a few ulps at most.
		const px = x <= 0.5 ? Math.pow(x, a) : Math.exp(a * Math.log1p(-y));
		const py = y <= 0.5 ? Math.pow(y, b) : Math.exp(b * Math.log1p(-x));
		if (px > 1e-280 && py > 1e-280) {
			const sum = a + b;
			const gammaSum = gamma(sum) * (1 + digammaApprox(sum) * sumError(a, b, sum));
			return (px * py) / (gamma(a) / gammaSum) / gamma(b);
		}
	}
	const lx = x < 0.5 ? Math.log(x) : Math.log1p(-y);
	const ly = y < 0.5 ? Math.log(y) : Math.log1p(-x);
	return Math.exp(a * lx + b * ly - lnBeta(a, b));
}

/**
 * I_x(a, b) by the continued fraction of DiDonato and Morris (TOMS 708,
 * BFRAC). It is written in terms of lambda = a - (a + b) x = (a + b) y - b,
 * which the caller forms from whichever of x and y is small, so the
 * fraction keeps full precision even when x is close to 1 and a >> b (the
 * Numerical Recipes form cancels catastrophically there).
 */
function betaFraction(a: number, b: number, x: number, y: number, lambda: number): number {
	const prefix = betaPrefix(a, b, x, y);
	if (prefix === 0) {
		return 0;
	}
	const c = lambda + 1;
	const c0 = b / a;
	const c1 = 1 / a + 1;
	const yp1 = y + 1;
	let p = 1;
	let s = a + 1;
	let an = 0;
	let bn = 1;
	let anp1 = 1;
	let bnp1 = c / c1;
	let r = c1 / c;
	// Convergence takes at most a few times 1e4 steps even for shapes near 1e12; the cap is a backstop.
	const limit = Math.min(1e6, 10000 + Math.ceil(20 * Math.sqrt(Math.max(a, b))));
	for (let n = 1; n <= limit; n++) {
		let t = n / a;
		const w = n * (b - n) * x;
		let e = a / s;
		const alpha = p * (p + c0) * e * e * (w * x);
		e = (t + 1) / (c1 + t + t);
		const beta = n + w / s + e * (c + n * yp1);
		p = t + 1;
		s += 2;
		t = alpha * an + beta * anp1;
		an = anp1;
		anp1 = t;
		t = alpha * bn + beta * bnp1;
		bn = bnp1;
		bnp1 = t;
		const r0 = r;
		r = anp1 / bnp1;
		if (Math.abs(r - r0) <= EPS * r) {
			break;
		}
		// Rescale to keep the recurrences in range.
		an /= bnp1;
		bn /= bnp1;
		anp1 = r;
		bnp1 = 1;
	}
	return prefix * r;
}

/** I_x(a, b) and its complement, for a, b > 0 and y = 1 - x. */
export function incBeta(a: number, b: number, x: number, y: number): TailPair {
	if (!(a > 0) || !(b > 0) || Number.isNaN(x) || Number.isNaN(y)) {
		return { lower: NaN, upper: NaN };
	}
	if (x <= 0) {
		return { lower: 0, upper: 1 };
	}
	if (y <= 0) {
		return { lower: 1, upper: 0 };
	}
	if (a === 1 && b === 1) {
		return { lower: x, upper: y };
	}
	if (b === 1) {
		// I_x(a, 1) = x^a
		const lx = x < 0.5 ? Math.log(x) : Math.log1p(-y);
		return { lower: x < 0.5 ? Math.pow(x, a) : Math.exp(a * lx), upper: -Math.expm1(a * lx) };
	}
	if (a === 1) {
		// I_x(1, b) = 1 - y^b
		const ly = y < 0.5 ? Math.log(y) : Math.log1p(-x);
		return { lower: -Math.expm1(b * ly), upper: y < 0.5 ? Math.pow(y, b) : Math.exp(b * ly) };
	}
	// Evaluate below the mean (lambda >= 0), where the fraction converges quickly.
	const lambda = x <= y ? a - (a + b) * x : (a + b) * y - b;
	if (lambda < 0) {
		const upper = betaFraction(b, a, y, x, -lambda);
		return { lower: 1 - upper, upper };
	}
	const lower = betaFraction(a, b, x, y, lambda);
	return { lower, upper: 1 - lower };
}

/** Density of the Beta(a, b) distribution at x (y = 1 - x). */
export function betaDensity(a: number, b: number, x: number, y: number): number {
	if (x <= 0 || y <= 0) {
		const atZero = x <= 0;
		const shape = atZero ? a : b;
		if (shape < 1) {
			return Infinity;
		}
		if (shape > 1) {
			return 0;
		}
		// Shape exactly 1: the density at the boundary is 1 / B(a, b).
		return a + b < 170 ? gamma(a + b) / gamma(a) / gamma(b) : Math.exp(-lnBeta(a, b));
	}
	return betaPrefix(a, b, x, y) / (x * y);
}

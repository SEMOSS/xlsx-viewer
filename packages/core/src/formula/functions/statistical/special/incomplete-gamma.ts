import { bd0, log1pmx } from './deviance';
import { erfc, expNegSquare } from './erf';
import { gamma, lnGamma, stirlerr } from './gamma-function';
import { normInv } from './normal';
import { solveQuantile } from './solve';

/**
 * Regularized incomplete gamma functions P(a, x) and Q(a, x) = 1 - P(a, x):
 * the power series for x < a + 1 and the Legendre continued fraction for Q
 * otherwise (Numerical Recipes 6.2). The prefactor x^a e^-x / Gamma(a) uses
 * Loader's saddle-point form for large a, so no precision is lost to large
 * cancelling logarithms. Beyond a = 1e8, where the series would need more
 * than 1e5 terms, Temme's uniform asymptotic expansion takes over.
 */

export interface TailPair {
	lower: number;
	upper: number;
}

const EPS = 1e-16;
const TINY = 1e-300;

/** Shape above which the uniform asymptotic expansion is used (its error is about 1e-3 / a). */
const TEMME_SHAPE = 1e8;

/** Taylor coefficients of Temme's C0(eta) = 1 / (lambda - 1) - 1 / eta (DiDonato and Morris). */
const TEMME_C0 = [
	-1 / 3,
	1 / 12,
	-2 / 135,
	1 / 864,
	1 / 2835,
	-139 / 777600,
	3.9192631785224378e-5,
	-2.1854485106799922e-6,
	-1.85406221071516e-6,
];

function iterationLimit(a: number): number {
	return 1000 + Math.ceil(12 * Math.sqrt(a));
}

/**
 * Temme's expansion Q(a, x) = erfc(eta sqrt(a / 2)) / 2 + e^(-a eta^2 / 2) / sqrt(2 pi a) C0(eta),
 * with lambda = x / a and eta = sign(lambda - 1) sqrt(2 (lambda - 1 - ln lambda)).
 */
function temmeLarge(a: number, x: number): TailPair {
	const mu = (x - a) / a;
	const eta = (mu < 0 ? -1 : 1) * Math.sqrt(-2 * log1pmx(mu));
	let c0: number;
	if (Math.abs(mu) < 0.05) {
		c0 = 0;
		for (let i = TEMME_C0.length - 1; i >= 0; i--) {
			c0 = c0 * eta + TEMME_C0[i];
		}
	} else {
		c0 = 1 / mu - 1 / eta;
	}
	const z = eta * Math.sqrt(a / 2);
	const r = (expNegSquare(z) / Math.sqrt(2 * Math.PI * a)) * c0;
	return { lower: 0.5 * erfc(-z) - r, upper: 0.5 * erfc(z) + r };
}

/** x^a e^(-x) / Gamma(a), the common prefactor of the incomplete gamma functions. */
export function gammaPrefix(a: number, x: number): number {
	if (x === 0) {
		return 0;
	}
	if (a >= 10) {
		return Math.sqrt(a / (2 * Math.PI)) * Math.exp(-bd0(a, x) - stirlerr(a));
	}
	if (x < 700) {
		const v = (Math.pow(x, a) * Math.exp(-x)) / gamma(a);
		if (v > 1e-280 && Number.isFinite(v)) {
			return v;
		}
	}
	return Math.exp(a * Math.log(x) - x - lnGamma(a));
}

function lowerSeries(a: number, x: number): number {
	let sum = 1;
	let term = 1;
	const limit = iterationLimit(a);
	for (let n = 1; n < limit; n++) {
		term *= x / (a + n);
		sum += term;
		if (term < sum * EPS) {
			break;
		}
	}
	return (gammaPrefix(a, x) / a) * sum;
}

function upperFraction(a: number, x: number): number {
	let b = x + 1 - a;
	let c = 1 / TINY;
	let d = 1 / b;
	let h = d;
	const limit = iterationLimit(a);
	for (let i = 1; i < limit; i++) {
		const an = -i * (i - a);
		b += 2;
		d = an * d + b;
		if (Math.abs(d) < TINY) {
			d = TINY;
		}
		c = b + an / c;
		if (Math.abs(c) < TINY) {
			c = TINY;
		}
		d = 1 / d;
		const delta = d * c;
		h *= delta;
		if (Math.abs(delta - 1) < EPS) {
			break;
		}
	}
	return gammaPrefix(a, x) * h;
}

/** P(a, x) and Q(a, x) for a > 0, x >= 0 (NaN for invalid arguments). */
export function incGamma(a: number, x: number): TailPair {
	if (!(a > 0) || !(x >= 0)) {
		return { lower: NaN, upper: NaN };
	}
	if (x === 0) {
		return { lower: 0, upper: 1 };
	}
	if (x === Infinity) {
		return { lower: 1, upper: 0 };
	}
	if (a === 1) {
		return { lower: -Math.expm1(-x), upper: Math.exp(-x) };
	}
	if (a > TEMME_SHAPE) {
		return temmeLarge(a, x);
	}
	if (x < a + 1) {
		const lower = lowerSeries(a, x);
		return { lower, upper: 1 - lower };
	}
	const upper = upperFraction(a, x);
	return { lower: 1 - upper, upper };
}

/** Density of the standard gamma distribution with shape a. */
export function gammaDensity(a: number, x: number): number {
	if (x === 0) {
		return a < 1 ? Infinity : a === 1 ? 1 : 0;
	}
	return gammaPrefix(a, x) / x;
}

/** Initial quantile estimate (Numerical Recipes 6.2.1 invgammp). */
function gammaGuess(a: number, p: number): number {
	if (a > 1) {
		const z = normInv(p);
		const t = 1 - 1 / (9 * a) + z / (3 * Math.sqrt(a));
		return Math.max(1e-3, a * t * t * t);
	}
	const t = 1 - a * (0.253 + a * 0.12);
	if (p < t) {
		return Math.pow(p / t, 1 / a);
	}
	return 1 - Math.log(1 - (p - t) / (1 - t));
}

/**
 * Quantile of the standard gamma distribution with shape a: x with
 * P(a, x) = prob, or Q(a, x) = prob when `upperTail` is set.
 */
export function gammaQuantile(a: number, prob: number, upperTail = false): number {
	const p = upperTail ? 1 - prob : prob;
	const q = upperTail ? prob : 1 - prob;
	if (p <= 0) {
		return 0;
	}
	if (q <= 0) {
		return Infinity;
	}
	const useLower = p <= 0.5;
	const guess = gammaGuess(a, p);
	return solveQuantile(
		{
			tail: (x) => (useLower ? incGamma(a, x).lower : incGamma(a, x).upper),
			density: (x) => gammaDensity(a, x),
			increasing: useLower,
		},
		useLower ? p : q,
		guess,
	);
}

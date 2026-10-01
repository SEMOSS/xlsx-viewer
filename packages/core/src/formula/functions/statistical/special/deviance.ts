import { lnGamma, stirlerr } from './gamma-function';

/**
 * Saddle-point building blocks (Loader, "Fast and Accurate Computation of
 * Binomial Probabilities", 2000): the deviance term bd0 and log1p(x) - x,
 * which let probability masses be formed without the cancellation of
 * subtracting large logarithms.
 */

const DBL_MIN = 2.2250738585072014e-308;

/** bd0(x, np) = x ln(x / np) + np - x, accurate when x is close to np. */
export function bd0(x: number, np: number): number {
	if (x === 0) {
		return np;
	}
	if (Math.abs(x - np) < 0.1 * (x + np)) {
		let v = (x - np) / (x + np);
		let s = (x - np) * v;
		if (Math.abs(s) < DBL_MIN) {
			return s;
		}
		let ej = 2 * x * v;
		v *= v;
		for (let j = 1; j < 1000; j++) {
			ej *= v;
			const s1 = s + ej / (2 * j + 1);
			if (s1 === s) {
				return s1;
			}
			s = s1;
		}
		return s;
	}
	return x * Math.log(x / np) + np - x;
}

/** log1p(x) - x without cancellation for small |x|. */
export function log1pmx(x: number): number {
	if (x > 1 || x < -0.79149064) {
		return Math.log1p(x) - x;
	}
	// log(1 + x) - x = r (2 y S(y) - x) with r = x / (2 + x), y = r^2, S(y) = sum y^k / (2k + 3).
	const r = x / (2 + x);
	const y = r * r;
	if (Math.abs(x) < 1e-2) {
		return r * (((((2 / 9) * y + 2 / 7) * y + 2 / 5) * y + 2 / 3) * y - x);
	}
	let sum = 0;
	let power = 1;
	for (let k = 0; k < 400; k++) {
		const term = power / (2 * k + 3);
		sum += term;
		if (term < 1e-17 * sum) {
			break;
		}
		power *= y;
	}
	return r * (2 * y * sum - x);
}

/** Poisson probability mass e^-lambda lambda^x / x! for real x >= 0. */
export function poissonPmf(x: number, lambda: number): number {
	if (lambda === 0) {
		return x === 0 ? 1 : 0;
	}
	if (x < 0 || !Number.isFinite(lambda)) {
		return 0;
	}
	if (x === 0) {
		return Math.exp(-lambda);
	}
	if (lambda < x * DBL_MIN) {
		return Math.exp(-lambda + x * Math.log(lambda) - lnGamma(x + 1));
	}
	return Math.exp(-stirlerr(x) - bd0(x, lambda)) / Math.sqrt(2 * Math.PI * x);
}

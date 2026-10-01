import { LN_SQRT_2PI, SQRT_2PI } from './erf';

/**
 * The gamma function and its logarithm.
 *
 * - x >= 10: Stirling's series with eight Bernoulli terms (error < 2e-18).
 * - 0.5 <= x <= 2.5: the zeta series of ln Gamma(2 + z) for |z| <= 0.5, which
 *   keeps full relative precision next to the zeros at x = 1 and x = 2.
 * - elsewhere below 10: the recurrence Gamma(x + 1) = x Gamma(x).
 * - x < 0: the reflection formula.
 */

/** 1 - Euler's constant. */
const ONE_MINUS_EULER = 0.42278433509846713;

/** (-1)^n (zeta(n) - 1) / n for n = 2 .. 30. */
const ZETA_SERIES = [
	0.3224670334241132, -0.0673523010531981, 0.020580808427784546, -0.007385551028673986,
	0.0028905103307415234, -0.001192753911703261, 0.0005096695247430425, -0.00022315475845357939,
	9.945751278180853e-5, -4.492623673813314e-5, 2.050721277567069e-5, -9.439488275268397e-6,
	4.374866789907488e-6, -2.039215753801366e-6, 9.55141213040742e-7, -4.492469198764566e-7,
	2.1207184805554665e-7, -1.0043224823968099e-7, 4.7698101693639804e-8, -2.2711094608943164e-8,
	1.0838659214896955e-8, -5.183475041970047e-9, 2.4836745438024785e-9, -1.1921401405860912e-9,
	5.731367241678862e-10, -2.7595228851242334e-10, 1.330476437424449e-10, -6.4229645638381e-11,
	3.1044247747322276e-11,
];

/** Loader's stirlerr(n) for n = 0, 0.5, 1, ..., 15 (index 2n). */
const STIRLERR_HALVES = [
	0, 0.15342640972002736, 0.08106146679532726, 0.05481412105191765, 0.0413406959554093,
	0.03316287351993629, 0.02767792568499834, 0.023746163656297496, 0.020790672103765093,
	0.018488450532673187, 0.016644691189821193, 0.015134973221917378, 0.013876128823070748,
	0.012810465242920227, 0.01189670994589177, 0.011104559758206917, 0.010411265261972096,
	0.009799416126158804, 0.009255462182712733, 0.008768700134139386, 0.00833056343336287,
	0.00793411456431402, 0.007573675487951841, 0.007244554301320383, 0.00694284010720953,
	0.006665247032707682, 0.006408994188004207, 0.006171712263039458, 0.0059513701127588475,
	0.0057462165130101155, 0.005554733551962801,
];

/** Exact factorials: FACTORIALS[n] = n! for n <= 22 (all exactly representable). */
const FACTORIALS: readonly number[] = (() => {
	const f = [1];
	for (let i = 1; i <= 22; i++) {
		f.push(f[i - 1] * i);
	}
	return f;
})();

/** Stirling-series correction ln Gamma(x) - [(x - 1/2) ln x - x + ln sqrt(2 pi)], for x >= 10. */
export function stirlingCorrection(x: number): number {
	const r = 1 / (x * x);
	return (
		(1 / 12 -
			r *
				(1 / 360 -
					r *
						(1 / 1260 -
							r *
								(1 / 1680 -
									r * (1 / 1188 - r * (691 / 360360 - r * (1 / 156 - (r * 3617) / 122400))))))) /
		x
	);
}

/** ln Gamma(2 + z) for |z| <= 0.5. */
function lnGamma2p(z: number): number {
	let poly = 0;
	for (let i = ZETA_SERIES.length - 1; i >= 0; i--) {
		poly = poly * z + ZETA_SERIES[i];
	}
	return z * ONE_MINUS_EULER + z * z * poly;
}

/** ln Gamma(x) for x > 0 (NaN otherwise). */
export function lnGamma(x: number): number {
	if (!(x > 0)) {
		return NaN;
	}
	if (x === Infinity) {
		return Infinity;
	}
	if (x >= 10) {
		return (x - 0.5) * Math.log(x) - x + LN_SQRT_2PI + stirlingCorrection(x);
	}
	if (x < 0.5) {
		return lnGamma(x + 1) - Math.log(x);
	}
	if (x < 1.5) {
		const z = x - 1;
		return lnGamma2p(z) - Math.log1p(z);
	}
	let z = x;
	let prod = 1;
	while (z > 2.5) {
		z -= 1;
		prod *= z;
	}
	return lnGamma2p(z - 2) + Math.log(prod);
}

/** sin(pi x), exactly zero at integers and accurate for large |x|. */
export function sinPi(x: number): number {
	if (Number.isInteger(x)) {
		return 0;
	}
	let r = x % 2;
	if (r < 0) {
		r += 2;
	}
	if (r <= 0.25) {
		return Math.sin(Math.PI * r);
	}
	if (r <= 0.75) {
		return Math.cos(Math.PI * (r - 0.5));
	}
	if (r <= 1.25) {
		return Math.sin(Math.PI * (1 - r));
	}
	if (r <= 1.75) {
		return -Math.cos(Math.PI * (r - 1.5));
	}
	return Math.sin(Math.PI * (r - 2));
}

/** The gamma function; NaN at the poles 0, -1, -2, ... and Infinity on overflow. */
export function gamma(x: number): number {
	if (Number.isNaN(x)) {
		return NaN;
	}
	if (Number.isInteger(x)) {
		if (x <= 0) {
			return NaN;
		}
		if (x <= 23) {
			return FACTORIALS[x - 1];
		}
	}
	if (x < 0) {
		return Math.PI / (sinPi(x) * gamma(1 - x));
	}
	if (x >= 10) {
		if (x > 171.7) {
			return Infinity;
		}
		// x^(x - 1/2) is split in two halves so neither overflows before e^-x is applied.
		const h = Math.pow(x, (x - 0.5) / 2);
		return SQRT_2PI * h * (h * Math.exp(-x)) * Math.exp(stirlingCorrection(x));
	}
	if (x < 0.5) {
		return gamma(x + 1) / x;
	}
	if (x < 1.5) {
		return Math.exp(lnGamma2p(x - 1)) / x;
	}
	let z = x;
	let prod = 1;
	while (z > 2.5) {
		z -= 1;
		prod *= z;
	}
	return Math.exp(lnGamma2p(z - 2)) * prod;
}

/**
 * Loader's stirlerr(n) = ln(n!) - ln(sqrt(2 pi n) (n / e)^n), the error of
 * Stirling's approximation. Exact table values at half integers up to 10.
 */
export function stirlerr(n: number): number {
	if (n >= 10) {
		return stirlingCorrection(n);
	}
	const twice = n + n;
	if (Number.isInteger(twice)) {
		return STIRLERR_HALVES[twice];
	}
	return lnGamma(n + 1) - (n + 0.5) * Math.log(n) + n - LN_SQRT_2PI;
}

/** ln Gamma(h + l) - ln Gamma(h) for h >= 10 and l >= 0, without cancellation. */
export function lnGammaRatio(h: number, l: number): number {
	return (
		l * Math.log(h) +
		(h + l - 0.5) * Math.log1p(l / h) -
		l +
		stirlingCorrection(h + l) -
		stirlingCorrection(h)
	);
}

/** n! for integer n >= 0 (exact up to 22!, Infinity above 170!). */
export function factorial(n: number): number {
	return n <= 22 ? FACTORIALS[n] : gamma(n + 1);
}

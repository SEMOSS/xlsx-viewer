/**
 * Bessel functions BESSELJ, BESSELY, BESSELI and BESSELK of integer order.
 *
 * Excel evaluates these with the classic rational / polynomial
 * approximations (Abramowitz & Stegun 9.4 / 9.8 as tabulated in Numerical
 * Recipes) plus recurrences, which are accurate to about 1e-8. They are
 * reproduced here on purpose so that results match Excel digit for digit
 * (BESSELJ(1.9, 2) is 0.329925829 in Excel; the exact value is 0.3299257277).
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite } from '../datetime/args';
import { engNumbers } from './args';

const TWO_OVER_PI = 0.636619772;

/** Evaluate `c[0] + y * (c[1] + y * (c[2] + ...))`. */
function poly(y: number, c: readonly number[]): number {
	let r = 0;
	for (let i = c.length - 1; i >= 0; i--) {
		r = r * y + c[i];
	}
	return r;
}

// Large-argument modulus / phase terms shared by J0, Y0 (P0, Q0) and J1, Y1 (P1, Q1).
const P0 = [1, -0.1098628627e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6];
const Q0_J = [
	-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934935152e-7,
];
const Q0_Y = [
	-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934945152e-7,
];
const P1 = [1, 0.183105e-2, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6];
const Q1 = [0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6];

function asymptotic(
	ax: number,
	phase: number,
	p: readonly number[],
	q: readonly number[],
	sine: boolean,
): number {
	const z = 8 / ax;
	const y = z * z;
	const xx = ax - phase;
	const a = poly(y, p);
	const b = z * poly(y, q);
	const amp = Math.sqrt(TWO_OVER_PI / ax);
	return sine
		? amp * (Math.sin(xx) * a + Math.cos(xx) * b)
		: amp * (Math.cos(xx) * a - Math.sin(xx) * b);
}

function j0(x: number): number {
	const ax = Math.abs(x);
	if (ax < 8) {
		const y = x * x;
		return (
			poly(
				y,
				[57568490574.0, -13362590354.0, 651619640.7, -11214424.18, 77392.33017, -184.9052456],
			) / poly(y, [57568490411.0, 1029532985.0, 9494680.718, 59272.64853, 267.8532712, 1.0])
		);
	}
	return asymptotic(ax, 0.785398164, P0, Q0_J, false);
}

function j1(x: number): number {
	const ax = Math.abs(x);
	if (ax < 8) {
		const y = x * x;
		return (
			(x *
				poly(
					y,
					[72362614232.0, -7895059235.0, 242396853.1, -2972611.439, 15704.4826, -30.16036606],
				)) /
			poly(y, [144725228442.0, 2300535178.0, 18583304.74, 99447.43394, 376.9991397, 1.0])
		);
	}
	const r = asymptotic(ax, 2.356194491, P1, Q1, false);
	return x < 0 ? -r : r;
}

function y0(x: number): number {
	if (x < 8) {
		const y = x * x;
		const r =
			poly(y, [-2957821389.0, 7062834065.0, -512359803.6, 10879881.29, -86327.92757, 228.4622733]) /
			poly(y, [40076544269.0, 745249964.8, 7189466.438, 47447.2647, 226.1030244, 1.0]);
		return r + TWO_OVER_PI * j0(x) * Math.log(x);
	}
	return asymptotic(x, 0.785398164, P0, Q0_Y, true);
}

function y1(x: number): number {
	if (x < 8) {
		const y = x * x;
		const r =
			(x *
				poly(
					y,
					[
						-0.4900604943e13, 0.127527439e13, -0.5153438139e11, 0.7349264551e9, -0.4237922726e7,
						0.8511937935e4,
					],
				)) /
			poly(
				y,
				[
					0.249958057e14, 0.4244419664e12, 0.3733650367e10, 0.2245904002e8, 0.102042605e6,
					0.3549632885e3, 1,
				],
			);
		return r + TWO_OVER_PI * (j1(x) * Math.log(x) - 1 / x);
	}
	return asymptotic(x, 2.356194491, P1, Q1, true);
}

function i0(x: number): number {
	const ax = Math.abs(x);
	if (ax < 3.75) {
		const y = (x / 3.75) ** 2;
		return poly(y, [1, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.360768e-1, 0.45813e-2]);
	}
	const y = 3.75 / ax;
	const c = [
		0.39894228, 0.1328592e-1, 0.225319e-2, -0.157565e-2, 0.916281e-2, -0.2057706e-1, 0.2635537e-1,
		-0.1647633e-1, 0.392377e-2,
	];
	return (Math.exp(ax) / Math.sqrt(ax)) * poly(y, c);
}

function i1(x: number): number {
	const ax = Math.abs(x);
	let r: number;
	if (ax < 3.75) {
		const y = (x / 3.75) ** 2;
		r =
			ax *
			poly(y, [0.5, 0.87890594, 0.51498869, 0.15084934, 0.2658733e-1, 0.301532e-2, 0.32411e-3]);
	} else {
		const y = 3.75 / ax;
		const tail = poly(y, [0.2282967e-1, -0.2895312e-1, 0.1787654e-1, -0.420059e-2]);
		r =
			poly(y, [0.39894228, -0.3988024e-1, -0.362018e-2, 0.163801e-2, -0.1031555e-1]) +
			y ** 5 * tail;
		r *= Math.exp(ax) / Math.sqrt(ax);
	}
	return x < 0 ? -r : r;
}

function k0(x: number): number {
	if (x <= 2) {
		const y = (x * x) / 4;
		return (
			-Math.log(x / 2) * i0(x) +
			poly(y, [-0.57721566, 0.4227842, 0.23069756, 0.348859e-1, 0.262698e-2, 0.1075e-3, 0.74e-5])
		);
	}
	const y = 2 / x;
	return (
		(Math.exp(-x) / Math.sqrt(x)) *
		poly(
			y,
			[
				1.25331414, -0.7832358e-1, 0.2189568e-1, -0.1062446e-1, 0.587872e-2, -0.25154e-2,
				0.53208e-3,
			],
		)
	);
}

function k1(x: number): number {
	if (x <= 2) {
		const y = (x * x) / 4;
		return (
			Math.log(x / 2) * i1(x) +
			(1 / x) *
				poly(y, [1, 0.15443144, -0.67278579, -0.18156897, -0.1919402e-1, -0.110404e-2, -0.4686e-4])
		);
	}
	const y = 2 / x;
	return (
		(Math.exp(-x) / Math.sqrt(x)) *
		poly(
			y,
			[1.25331414, 0.23498619, -0.365562e-1, 0.1504268e-1, -0.780353e-2, 0.325614e-2, -0.68245e-3],
		)
	);
}

/** Forward recurrence `f(k+1) = (2k/x) f(k) + sign * f(k-1)` from orders 0 and 1. */
function forward(x: number, n: number, f0: number, f1: number, sign: 1 | -1): number {
	if (n === 0) {
		return f0;
	}
	let prev = f0;
	let cur = f1;
	for (let k = 1; k < n; k++) {
		const next = ((2 * k) / x) * cur + sign * prev;
		prev = cur;
		cur = next;
	}
	return cur;
}

/** Miller's downward recurrence start index for order `n` (Numerical Recipes, ACC = 40). */
function millerStart(n: number): number {
	return 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2);
}

export function besselJ(x: number, n: number): number {
	if (n < 2) {
		return n === 0 ? j0(x) : j1(x);
	}
	const ax = Math.abs(x);
	if (ax === 0) {
		return 0;
	}
	let ans: number;
	if (ax > n) {
		ans = forward(ax, n, j0(ax), j1(ax), -1);
	} else {
		// Downward recurrence normalised with J0 + 2 (J2 + J4 + ...) = 1.
		const tox = 2 / ax;
		let even = false;
		let bjp = 0;
		let bj = 1;
		let sum = 0;
		ans = 0;
		for (let j = millerStart(n); j > 0; j--) {
			const bjm = j * tox * bj - bjp;
			bjp = bj;
			bj = bjm;
			if (Math.abs(bj) > 1e10) {
				bj *= 1e-10;
				bjp *= 1e-10;
				ans *= 1e-10;
				sum *= 1e-10;
			}
			if (even) {
				sum += bj;
			}
			even = !even;
			if (j === n) {
				ans = bjp;
			}
		}
		ans /= 2 * sum - bj;
	}
	return x < 0 && n % 2 === 1 ? -ans : ans;
}

export function besselI(x: number, n: number): number {
	if (n < 2) {
		return n === 0 ? i0(x) : i1(x);
	}
	if (x === 0) {
		return 0;
	}
	const tox = 2 / Math.abs(x);
	let bip = 0;
	let bi = 1;
	let ans = 0;
	for (let j = 2 * (n + Math.floor(Math.sqrt(40 * n))); j > 0; j--) {
		const bim = bip + j * tox * bi;
		bip = bi;
		bi = bim;
		if (Math.abs(bi) > 1e10) {
			ans *= 1e-10;
			bi *= 1e-10;
			bip *= 1e-10;
		}
		if (j === n) {
			ans = bip;
		}
	}
	ans *= i0(x) / bi;
	return x < 0 && n % 2 === 1 ? -ans : ans;
}

export function besselY(x: number, n: number): number {
	return forward(x, n, y0(x), y1(x), -1);
}

export function besselK(x: number, n: number): number {
	return forward(x, n, k0(x), k1(x), 1);
}

/** `x`, `n` with n truncated; n < 0 (and x <= 0 for Y / K) is #NUM!. */
function bessel(f: (x: number, n: number) => number, positiveX: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const v = engNumbers(args, ctx, [null, null]);
		if (isCellError(v)) {
			return v;
		}
		const x = v[0];
		const n = Math.trunc(v[1]);
		if (n < 0 || (positiveX && x <= 0)) {
			return ERRORS.NUM;
		}
		return finite(f(x, n));
	};
}

const def = (name: string, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: 2,
	maxArgs: 2,
	params: ['value'],
	impl,
});

export const BESSEL_FUNCTIONS: FunctionDefinition[] = [
	def('BESSELI', bessel(besselI, false)),
	def('BESSELJ', bessel(besselJ, false)),
	def('BESSELK', bessel(besselK, true)),
	def('BESSELY', bessel(besselY, true)),
];

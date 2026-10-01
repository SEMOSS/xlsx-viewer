import { erf, erfcx, expNegHalfSquare, SQRT_2PI } from './erf';

/** Standard normal density. */
export function normPdf(z: number): number {
	return expNegHalfSquare(z) / SQRT_2PI;
}

/** Standard normal upper tail P(Z > z), accurate far into the tail. */
export function normSf(z: number): number {
	if (Number.isNaN(z)) {
		return NaN;
	}
	if (z < 0) {
		return 1 - normSf(-z);
	}
	const x = z * Math.SQRT1_2;
	if (x < 0.75) {
		return 0.5 * (1 - erf(x));
	}
	return 0.5 * erfcx(x) * expNegHalfSquare(z);
}

/** Standard normal distribution function P(Z <= z). */
export function normCdf(z: number): number {
	return normSf(-z);
}

function polyval(coefficients: readonly number[], x: number): number {
	let r = 0;
	for (let i = coefficients.length - 1; i >= 0; i--) {
		r = r * x + coefficients[i];
	}
	return r;
}

// Wichura (1988), algorithm AS 241 (PPND16), coefficients in ascending powers.
const A = [
	3.387132872796366608, 1.3314166789178437745e2, 1.9715909503065514427e3, 1.3731693765509461125e4,
	4.5921953931549871457e4, 6.7265770927008700853e4, 3.3430575583588128105e4,
	2.5090809287301226727e3,
];
const B = [
	1, 4.2313330701600911252e1, 6.871870074920579083e2, 5.3941960214247511077e3,
	2.1213794301586595867e4, 3.930789580009271061e4, 2.8729085735721942674e4, 5.226495278852854561e3,
];
const C = [
	1.42343711074968357734, 4.6303378461565452959, 5.7694972214606914055, 3.64784832476320460504,
	1.27045825245236838258, 2.4178072517745061177e-1, 2.27238449892691845833e-2,
	7.7454501427834140764e-4,
];
const D = [
	1, 2.05319162663775882187, 1.6763848301838038494, 6.8976733498510000455e-1,
	1.4810397642748007459e-1, 1.51986665636164571966e-2, 5.475938084995344946e-4,
	1.05075007164441684324e-9,
];
const E = [
	6.6579046435011037772, 5.4637849111641143699, 1.7848265399172913358, 2.9656057182850489123e-1,
	2.6532189526576123093e-2, 1.2426609473880784386e-3, 2.71155556874348757815e-5,
	2.01033439929228813265e-7,
];
const F = [
	1, 5.9983220655588793769e-1, 1.3692988092273580531e-1, 1.48753612908506148525e-2,
	7.868691311456132591e-4, 1.8463183175100546818e-5, 1.4215117583164458887e-7,
	2.04426310338993978564e-15,
];

/**
 * Standard normal quantile: z with P(Z <= z) = p, for 0 < p < 1. Wichura's
 * AS 241 is accurate to about 1e-16 relative over the whole range.
 */
export function normInv(p: number): number {
	if (!(p > 0 && p < 1)) {
		return p === 0 ? -Infinity : p === 1 ? Infinity : NaN;
	}
	const q = p - 0.5;
	if (Math.abs(q) <= 0.425) {
		const r = 0.180625 - q * q;
		return (q * polyval(A, r)) / polyval(B, r);
	}
	let r = Math.sqrt(-Math.log(q < 0 ? p : 1 - p));
	let v: number;
	if (r <= 5) {
		r -= 1.6;
		v = polyval(C, r) / polyval(D, r);
	} else {
		r -= 5;
		v = polyval(E, r) / polyval(F, r);
	}
	return q < 0 ? -v : v;
}

/** Standard normal upper quantile: z with P(Z > z) = q, accurate for tiny q. */
export function normInvUpper(q: number): number {
	return -normInv(q);
}

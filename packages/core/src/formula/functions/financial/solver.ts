/**
 * Root finding for the iterative financial functions (RATE, IRR, XIRR,
 * YIELD, ODDFYIELD).
 */

export interface NewtonOptions {
	/** Iterations allowed to reach `tolerance`. */
	maxIterations?: number;
	/** Convergence threshold on the step size (absolute). */
	tolerance?: number;
	/** Extra iterations run after convergence to polish the last digits. */
	polish?: number;
}

/**
 * Newton-Raphson iteration from `guess`. `fdf(x)` returns `[f(x), f'(x)]`.
 * Returns the root, or `null` when the iteration fails to converge within
 * `maxIterations` or leaves the domain (non-finite values).
 */
export function newton(
	fdf: (x: number) => [number, number],
	guess: number,
	{ maxIterations = 50, tolerance = 1e-10, polish = 3 }: NewtonOptions = {},
): number | null {
	let x = guess;
	for (let i = 0; i < maxIterations; i++) {
		const [fx, dfx] = fdf(x);
		if (!Number.isFinite(fx) || !Number.isFinite(dfx)) {
			return null;
		}
		if (fx === 0) {
			return x;
		}
		// Move away from a flat spot rather than dividing by zero.
		const next = dfx === 0 ? x + 1.1e-7 : x - fx / dfx;
		if (!Number.isFinite(next)) {
			return null;
		}
		const converged = Math.abs(next - x) < tolerance;
		x = next;
		if (converged) {
			return polishRoot(fdf, x, polish);
		}
	}
	return null;
}

function polishRoot(fdf: (x: number) => [number, number], start: number, steps: number): number {
	let x = start;
	let [fx] = fdf(x);
	for (let i = 0; i < steps && fx !== 0; i++) {
		const [, dfx] = fdf(x);
		if (dfx === 0 || !Number.isFinite(dfx)) {
			break;
		}
		const next = x - fx / dfx;
		const [fNext] = fdf(next);
		if (!Number.isFinite(fNext) || Math.abs(fNext) >= Math.abs(fx)) {
			break;
		}
		x = next;
		fx = fNext;
	}
	return x;
}

/** Bisection on a bracket `[lo, hi]` whose ends have opposite signs. */
export function bisect(
	f: (x: number) => number,
	lo: number,
	hi: number,
	iterations = 200,
): number | null {
	let a = lo;
	let b = hi;
	let fa = f(a);
	const fb = f(b);
	if (!Number.isFinite(fa) || !Number.isFinite(fb)) {
		return null;
	}
	if (fa === 0 || fb === 0) {
		return fa === 0 ? a : b;
	}
	if (Math.sign(fa) === Math.sign(fb)) {
		return null;
	}
	for (let i = 0; i < iterations; i++) {
		const m = (a + b) / 2;
		const fm = f(m);
		if (fm === 0 || m === a || m === b) {
			return m;
		}
		if (Math.sign(fm) === Math.sign(fa)) {
			a = m;
			fa = fm;
		} else {
			b = m;
		}
	}
	return (a + b) / 2;
}

/** Candidate points used to look for sign changes when Newton fails. */
const SCAN_POINTS = [
	-0.999999, -0.99, -0.95, -0.9, -0.8, -0.7, -0.6, -0.5, -0.4, -0.3, -0.2, -0.15, -0.1, -0.05,
	-0.02, 0, 0.02, 0.05, 0.1, 0.15, 0.2, 0.3, 0.4, 0.5, 0.75, 1, 1.5, 2, 3, 5, 10, 20, 50, 100, 1000,
	1e4, 1e5,
];

/**
 * Fallback for rate solvers on `(-1, inf)`: find the sign change nearest to
 * `guess` among a grid of rates, bisect it, then polish with Newton.
 */
export function scanForRoot(fdf: (x: number) => [number, number], guess: number): number | null {
	const f = (x: number) => fdf(x)[0];
	const points = [...new Set([...SCAN_POINTS, guess])].filter((p) => p > -1).sort((a, b) => a - b);
	let best: number | null = null;
	for (let i = 0; i + 1 < points.length; i++) {
		const lo = points[i];
		const hi = points[i + 1];
		const flo = f(lo);
		const fhi = f(hi);
		if (!Number.isFinite(flo) || !Number.isFinite(fhi) || Math.sign(flo) === Math.sign(fhi)) {
			continue;
		}
		const root = bisect(f, lo, hi);
		if (root !== null && (best === null || Math.abs(root - guess) < Math.abs(best - guess))) {
			best = root;
		}
	}
	return best === null ? null : polishRoot(fdf, best, 5);
}

import { betaDensity, incBeta } from './incomplete-beta';
import { solveQuantile } from './solve';

/** A quantile x of a distribution on (0, 1) together with y = 1 - x, both at full precision. */
export interface UnitQuantile {
	x: number;
	y: number;
}

/** Initial estimate of the Beta(a, b) quantile at lower-tail probability p (Numerical Recipes 6.14). */
function betaGuess(a: number, b: number, p: number): number {
	if (a >= 1 && b >= 1) {
		const pp = p < 0.5 ? p : 1 - p;
		const t = Math.sqrt(-2 * Math.log(pp));
		let z = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t;
		if (p < 0.5) {
			z = -z;
		}
		const al = (z * z - 3) / 6;
		const h = 2 / (1 / (2 * a - 1) + 1 / (2 * b - 1));
		const w =
			(z * Math.sqrt(al + h)) / h -
			(1 / (2 * b - 1) - 1 / (2 * a - 1)) * (al + 5 / 6 - 2 / (3 * h));
		return a / (a + b * Math.exp(2 * w));
	}
	const t = Math.exp(a * Math.log(a / (a + b))) / a;
	const u = Math.exp(b * Math.log(b / (a + b))) / b;
	const w = t + u;
	if (p < t / w) {
		return Math.pow(a * w * p, 1 / a);
	}
	return 1 - Math.pow(b * w * (1 - p), 1 / b);
}

/**
 * Quantile of the Beta(a, b) distribution at `prob` (a lower-tail
 * probability, or an upper-tail one when `upperTail` is set). The solver
 * works on whichever of x and 1 - x lies below 1/2 and targets whichever
 * tail probability is below 1/2, so both tails keep full precision.
 */
export function betaQuantile(a: number, b: number, prob: number, upperTail = false): UnitQuantile {
	const p = upperTail ? 1 - prob : prob;
	const q = upperTail ? prob : 1 - prob;
	if (p <= 0) {
		return { x: 0, y: 1 };
	}
	if (q <= 0) {
		return { x: 1, y: 0 };
	}
	const useLower = p <= 0.5;
	const target = useLower ? p : q;
	const mid = incBeta(a, b, 0.5, 0.5);
	const quantileBelowHalf = useLower ? mid.lower >= p : mid.upper <= q;
	if (quantileBelowHalf) {
		const x = solveQuantile(
			{
				tail: (v) => {
					const r = incBeta(a, b, v, 1 - v);
					return useLower ? r.lower : r.upper;
				},
				density: (v) => betaDensity(a, b, v, 1 - v),
				increasing: useLower,
				upper: 1,
			},
			target,
			betaGuess(a, b, p),
		);
		return { x, y: 1 - x };
	}
	// Solve for y = 1 - x, which follows Beta(b, a): P(X <= x) = P(Y >= y).
	const y = solveQuantile(
		{
			tail: (v) => {
				const r = incBeta(b, a, v, 1 - v);
				return useLower ? r.upper : r.lower;
			},
			density: (v) => betaDensity(b, a, v, 1 - v),
			increasing: !useLower,
			upper: 1,
		},
		target,
		betaGuess(b, a, q),
	);
	return { x: 1 - y, y };
}

/**
 * Safeguarded Newton iteration shared by the continuous quantile functions.
 *
 * Solves `tail(x) = target` for x in (0, upper), where `tail` is a monotone
 * probability (a distribution function or a survival function) and
 * `density` is |d tail / dx|. Newton steps are taken on ln(tail) against
 * ln(x), which is close to linear in both tails, and every evaluation
 * tightens a bracket around the root; a step that leaves the bracket is
 * replaced by bisection in ln(x), so the iteration always converges.
 */
export interface QuantileProblem {
	tail: (x: number) => number;
	density: (x: number) => number;
	/** `tail` grows with x (a distribution function) rather than shrinking. */
	increasing: boolean;
	/** Exclusive upper bound of the variable (default Infinity). */
	upper?: number;
}

const MAX_ITERATIONS = 300;
const RELATIVE_TOLERANCE = 1e-15;

/** Next probe when the bracket is open on one side, or its log midpoint. */
function bisect(lo: number, hi: number): number {
	if (hi === Infinity) {
		return lo >= 1 ? lo * lo + 1 : lo >= 0.1 ? lo * 10 : Math.sqrt(lo);
	}
	if (lo === 0) {
		// Shrink towards zero; the product underflows to 0 when the root is below the double range.
		return hi > 1 ? (hi >= 10 ? Math.sqrt(hi) : hi / 10) : hi * Math.min(hi, 0.5);
	}
	return Math.sqrt(lo) * Math.sqrt(hi);
}

/** Root of `tail(x) = target`, starting from `guess`. */
export function solveQuantile(problem: QuantileProblem, target: number, guess: number): number {
	const { tail, density, increasing } = problem;
	const upper = problem.upper ?? Infinity;
	let lo = 0;
	let hi = upper;
	let x = guess > 0 && guess < upper && Number.isFinite(guess) ? guess : Math.min(1, upper / 2);
	for (let iter = 0; iter < MAX_ITERATIONS; iter++) {
		const t = tail(x);
		if (t === target) {
			return x;
		}
		if (t < target === increasing) {
			lo = x;
		} else {
			hi = x;
		}
		let next = NaN;
		const f = density(x);
		if (t > 0 && f > 0 && Number.isFinite(f)) {
			const ratio = t / target;
			const h =
				ratio > 0 && Number.isFinite(ratio) ? Math.log(ratio) : Math.log(t) - Math.log(target);
			const slope = ((increasing ? 1 : -1) * x * f) / t;
			next = x * Math.exp(-h / slope);
		}
		if (!(next > lo && next < hi)) {
			next = bisect(lo, hi);
			if (next === 0) {
				return 0;
			}
		}
		if (Math.abs(next - x) <= RELATIVE_TOLERANCE * x) {
			return next;
		}
		x = next;
	}
	return x;
}

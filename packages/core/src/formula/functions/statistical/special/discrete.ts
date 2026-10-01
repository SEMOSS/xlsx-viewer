import { poissonPmf } from './deviance';
import { incGamma } from './incomplete-gamma';
import { betaPrefix, incBeta } from './incomplete-beta';
import { normInv } from './normal';

/**
 * Discrete distributions. Masses use the saddle-point form of the beta
 * prefactor; distribution functions sum masses for small supports and fall
 * back to the incomplete beta / gamma functions otherwise.
 */

/** Largest number of terms summed directly before switching to a special function. */
const SUM_LIMIT = 1000;

/** Rows of Pascal's triangle are exact in doubles up to n = 56 (C(56, 28) < 2^53). */
const PASCAL_LIMIT = 56;
let pascal: number[][] | undefined;

/** Exact binomial coefficient C(n, k) for 0 <= k <= n <= 56. */
export function exactBinomial(n: number, k: number): number {
	if (!pascal) {
		pascal = [[1]];
		for (let i = 1; i <= PASCAL_LIMIT; i++) {
			const prev = pascal[i - 1];
			const row = [1];
			for (let j = 1; j < i; j++) {
				row.push(prev[j - 1] + prev[j]);
			}
			row.push(1);
			pascal.push(row);
		}
	}
	return pascal[n][k];
}

/** Binomial mass C(n, k) p^k q^(n - k) for integers 0 <= k <= n; q = 1 - p. */
export function binomPmf(k: number, n: number, p: number, q: number): number {
	if (k < 0 || k > n) {
		return 0;
	}
	if (p === 0) {
		return k === 0 ? 1 : 0;
	}
	if (q === 0) {
		return k === n ? 1 : 0;
	}
	if (k === 0) {
		return Math.exp(n * (p < 0.5 ? Math.log1p(-p) : Math.log(q)));
	}
	if (k === n) {
		return Math.exp(n * (q < 0.5 ? Math.log1p(-q) : Math.log(p)));
	}
	if (n <= PASCAL_LIMIT) {
		// Exact coefficient and powers: textbook cases such as p = 1/2 come out exact.
		const pk = Math.pow(p, k);
		const qk = Math.pow(q, n - k);
		if (pk > 1e-280 && qk > 1e-280) {
			return exactBinomial(n, k) * pk * qk;
		}
	}
	// C(n, k) = n / (k (n - k)) / B(k, n - k)
	return (n / (k * (n - k))) * betaPrefix(k, n - k, p, q);
}

/** P(X <= k) for X ~ Binomial(n, p). */
export function binomCdf(k: number, n: number, p: number, q: number): number {
	if (k < 0) {
		return 0;
	}
	if (k >= n) {
		return 1;
	}
	if (n <= SUM_LIMIT) {
		let sum = 0;
		if (k < n * p) {
			for (let j = 0; j <= k; j++) {
				sum += binomPmf(j, n, p, q);
			}
			return sum;
		}
		for (let j = k + 1; j <= n; j++) {
			sum += binomPmf(j, n, p, q);
		}
		return 1 - sum;
	}
	return incBeta(n - k, k + 1, q, p).lower;
}

/** P(lo <= X <= hi) for X ~ Binomial(n, p), without cancellation for short ranges. */
export function binomRange(lo: number, hi: number, n: number, p: number, q: number): number {
	if (hi - lo <= SUM_LIMIT) {
		let sum = 0;
		for (let j = lo; j <= hi; j++) {
			sum += binomPmf(j, n, p, q);
		}
		return sum;
	}
	if (lo > n * p) {
		return 1 - binomCdf(lo - 1, n, p, q) - (1 - binomCdf(hi, n, p, q));
	}
	return binomCdf(hi, n, p, q) - binomCdf(lo - 1, n, p, q);
}

/** Smallest k with P(X <= k) >= alpha, X ~ Binomial(n, p). */
export function binomQuantile(n: number, p: number, alpha: number): number {
	const q = 1 - p;
	if (alpha <= 0 || p === 0) {
		return 0;
	}
	if (p === 1) {
		return n;
	}
	// Accept cumulative values within rounding of alpha.
	const target = alpha * (1 - 1e-13);
	const mu = n * p;
	const sigma = Math.sqrt(n * p * q);
	const z = alpha >= 1 ? 40 : normInv(alpha);
	let k = Math.max(0, Math.min(n, Math.floor(mu + sigma * z)));
	let c = binomCdf(k, n, p, q);
	if (c >= target) {
		while (k > 0) {
			const previous = c - binomPmf(k, n, p, q);
			if (previous < target) {
				break;
			}
			c = previous;
			k--;
		}
		return k;
	}
	while (k < n && c < target) {
		k++;
		c += binomPmf(k, n, p, q);
	}
	return k;
}

/** Negative binomial mass: f failures before the s-th success, success probability p. */
export function negBinomPmf(f: number, s: number, p: number, q: number): number {
	if (f < 0) {
		return 0;
	}
	if (f === 0) {
		return Math.exp(s * (q < 0.5 ? Math.log1p(-q) : Math.log(p)));
	}
	return (s / (f + s)) * binomPmf(s, f + s, p, q);
}

/** P(F <= f) for the negative binomial distribution. */
export function negBinomCdf(f: number, s: number, p: number, q: number): number {
	if (f < 0) {
		return 0;
	}
	return incBeta(s, f + 1, p, q).lower;
}

/** P(X <= k) for X ~ Poisson(lambda). */
export function poissonCdf(k: number, lambda: number): number {
	if (k < 0) {
		return 0;
	}
	if (lambda === 0) {
		return 1;
	}
	if (k <= SUM_LIMIT) {
		let sum = 0;
		for (let j = 0; j <= k; j++) {
			sum += poissonPmf(j, lambda);
		}
		return Math.min(1, sum);
	}
	return incGamma(k + 1, lambda).upper;
}

/**
 * Hypergeometric mass: x successes in a sample of n drawn without
 * replacement from a population of size total containing m successes.
 */
export function hypergeomPmf(x: number, n: number, m: number, total: number): number {
	if (x < 0 || x > n || x > m || n - x > total - m) {
		return 0;
	}
	if (n === 0) {
		return x === 0 ? 1 : 0;
	}
	if (total <= PASCAL_LIMIT) {
		return (exactBinomial(m, x) * exactBinomial(total - m, n - x)) / exactBinomial(total, n);
	}
	const p = n / total;
	const q = (total - n) / total;
	return (binomPmf(x, m, p, q) * binomPmf(n - x, total - m, p, q)) / binomPmf(n, total, p, q);
}

/** P(X <= x) for the hypergeometric distribution (R's phyper algorithm). */
export function hypergeomCdf(x: number, n: number, m: number, total: number): number {
	let red = m;
	let black = total - m;
	let k = x;
	let lower = true;
	if (k * total > n * red) {
		// Sum the shorter tail.
		const swap = black;
		black = red;
		red = swap;
		k = n - k - 1;
		lower = false;
	}
	if (k < 0 || k < n - black) {
		return lower ? 0 : 1;
	}
	if (k >= red || k >= n) {
		return lower ? 1 : 0;
	}
	const d = hypergeomPmf(k, n, red, red + black);
	let sum = 0;
	let term = 1;
	let j = k;
	while (j > 0 && term >= 1e-16 * sum) {
		term *= (j * (black - n + j)) / (n + 1 - j) / (red + 1 - j);
		sum += term;
		j--;
	}
	const tail = d * (1 + sum);
	return lower ? tail : 1 - tail;
}

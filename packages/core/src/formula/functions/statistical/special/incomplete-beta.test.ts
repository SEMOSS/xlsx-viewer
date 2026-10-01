import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { betaQuantile } from './beta-quantile';
import { lnGamma } from './gamma-function';
import { betaDensity, betaPrefix, incBeta, lnBeta } from './incomplete-beta';

// [a, b, x, I_x(a, b), 1 - I_x(a, b)] from 120-digit Decimal arithmetic.
const TABLE: [number, number, number, number, number][] = [
	[0.1, 25, 0.01, 0.8946276641300885, 0.10537233586991142],
	[0.5, 0.5, 0.5, 0.5, 0.5],
	[0.5, 1000, 0.0015497251374312843, 0.9217595853621472, 0.0782404146378528],
	[3.5, 1000, 0.004836571998006976, 0.79431055499639, 0.20568944500360997],
	[10, 10, 0.45, 0.3289640875783923, 0.6710359124216078],
	[25, 100, 0.5, 0.9999999999984034, 1.5966605690176048e-12],
	[100, 1000, 0.101, 0.8761524814284406, 0.12384751857155943],
	[1000, 0.1, 0.999, 0.024125403119185585, 0.9758745968808145],
	[1000, 1000, 0.5, 0.5, 0.5],
	[50000, 50000, 0.499, 0.26354493829704145, 0.7364550617029586],
	[200000, 10, 0.99995, 0.4578702871165965, 0.5421297128834035],
	[0.5, 500000, 1e-6, 0.6826894921371464, 0.31731050786285364],
	[30, 0.5, 0.97, 0.17821754497024134, 0.8217824550297587],
];

describe('incBeta', () => {
	it.each(TABLE)('a = %d, b = %d, x = %d', (a, b, x, lower, upper) => {
		const r = incBeta(a, b, x, 1 - x);
		expectClose(r.lower, lower, 1e-13);
		expectClose(r.upper, upper, 1e-13);
	});

	it('keeps precision for a huge first shape next to x = 1', () => {
		// P(|T| > t) for Student t with 1e10 degrees of freedom at the normal 5% point.
		const df = 1e10;
		const t = 1.95996398477726;
		const y = (t * t) / (df + t * t);
		expectClose(incBeta(df / 2, 0.5, 1 - y, y).lower, 0.0500000000000025, 1e-13);
	});

	it('uses the closed forms for a = 1 or b = 1', () => {
		expectClose(incBeta(3, 1, 0.5, 0.5).lower, 0.125, 1e-15);
		expectClose(incBeta(1, 4, 0.5, 0.5).upper, 0.0625, 1e-15);
		expect(incBeta(1, 1, 0.3, 0.7)).toEqual({ lower: 0.3, upper: 0.7 });
	});

	it('handles the boundaries and invalid shapes', () => {
		expect(incBeta(2, 3, 0, 1)).toEqual({ lower: 0, upper: 1 });
		expect(incBeta(2, 3, 1, 0)).toEqual({ lower: 1, upper: 0 });
		expect(incBeta(0, 3, 0.5, 0.5).lower).toBeNaN();
		expect(incBeta(2, -1, 0.5, 0.5).lower).toBeNaN();
	});

	it('computes ln B(a, b) in every regime', () => {
		expectClose(lnBeta(2, 3), Math.log(1 / 12), 1e-15);
		expectClose(lnBeta(0.5, 0.5), Math.log(Math.PI), 1e-15);
		expectClose(lnBeta(3, 12), lnGamma(3) + lnGamma(12) - lnGamma(15), 1e-14);
		expectClose(lnBeta(25, 40), lnGamma(25) + lnGamma(40) - lnGamma(65), 1e-14);
		// ln B(1/2, n) ~ ln Gamma(1/2) - ln(n) / 2 for large n
		expectClose(lnBeta(0.5, 1e12), lnGamma(0.5) - 0.5 * Math.log(1e12), 1e-12);
	});

	it('forms densities and prefactors', () => {
		expectClose(betaDensity(8, 10, 0.5, 0.5), 2.967529296875, 1e-14);
		expectClose(betaDensity(2, 5, 0.9, 0.1), 0.0027, 1e-13);
		expectClose(betaPrefix(20, 25, 0.42, 0.58), 0.42 * 0.58 * 5.138157576488628, 1e-13);
		expect(betaDensity(0.5, 2, 0, 1)).toBe(Infinity);
		expect(betaDensity(3, 2, 0, 1)).toBe(0);
		expectClose(betaDensity(1, 4, 0, 1), 4, 1e-15);
	});
});

describe('betaQuantile', () => {
	it.each([
		[0.5, 2, 3, 0.38572756813238956],
		[0.001, 0.5, 4, 2.0897967918172725e-7],
		[0.99, 10, 1.5, 0.9944112827578341],
		[0.3, 30, 40, 0.39698703679444164],
	])('p = %d, a = %d, b = %d', (p, a, b, x) => {
		const q = betaQuantile(a, b, p);
		expectClose(q.x, x, 1e-14);
		expectClose(q.y, 1 - x, 1e-12);
	});

	it('round-trips both tails, solving in whichever of x and 1 - x is small', () => {
		for (const [a, b] of [
			[0.5, 0.5],
			[2, 3],
			[0.1, 5],
			[5, 0.1],
			[20, 30],
			[200, 2],
			[1e4, 1e4],
		]) {
			for (const p of [1e-8, 0.01, 0.5, 0.9]) {
				const lo = betaQuantile(a, b, p);
				expectClose(incBeta(a, b, lo.x, lo.y).lower, p, 1e-11);
				const up = betaQuantile(a, b, p, true);
				expectClose(incBeta(a, b, up.x, up.y).upper, p, 1e-11);
			}
		}
	});

	it('handles the boundaries', () => {
		expect(betaQuantile(2, 3, 0)).toEqual({ x: 0, y: 1 });
		expect(betaQuantile(2, 3, 1)).toEqual({ x: 1, y: 0 });
		expect(betaQuantile(2, 3, 0, true)).toEqual({ x: 1, y: 0 });
	});
});

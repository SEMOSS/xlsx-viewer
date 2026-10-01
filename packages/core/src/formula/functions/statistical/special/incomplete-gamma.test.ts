import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { gammaDensity, gammaPrefix, gammaQuantile, incGamma } from './incomplete-gamma';

// [a, x, P(a, x), Q(a, x)] from 120-digit Decimal arithmetic.
const TABLE: [number, number, number, number][] = [
	[0.1, 0.0001, 0.41846137523796295, 0.581538624762037],
	[0.5, 0.05, 0.24817036595415073, 0.7518296340458492],
	[1.5, 1.5, 0.608374823728911, 0.39162517627108895],
	[2.5, 0.25, 0.007876706767370407, 0.9921232932326296],
	[5, 4.95, 0.5506899252403475, 0.4493100747596524],
	[9.5, 10.450000000000001, 0.6576473509359595, 0.34235264906404045],
	[10, 1, 1.1142547833872067e-7, 0.9999998885745217],
	[12.3, 61.5, 0.9999999999999953, 4.681707951484865e-15],
	[15.5, 15.5, 0.5337874937824917, 0.46621250621750837],
	[50, 49.5, 0.49052780120162587, 0.5094721987983741],
	[200, 220.00000000000003, 0.9181943116110619, 0.08180568838893805],
	[1000, 500, 3.2982727970671e-86, 1],
	[1000, 1010, 0.6276789447369947, 0.3723210552630053],
	[0.5, 30, 0.9999999999999906, 9.485737571073848e-15],
	[3, 100, 1, 1.8976107553682285e-40],
	[30, 1, 1.4330814167223182e-33, 1],
];

describe('incGamma', () => {
	it.each(TABLE)('a = %d, x = %d', (a, x, p, q) => {
		const r = incGamma(a, x);
		expectClose(r.lower, p, 1e-13);
		expectClose(r.upper, q, 1e-13);
	});

	it.each([
		[1e9, 1e9, 0.5000042052208701, 0.4999957947791299],
		[1e9, 1e9 + 50000, 0.9430750438329841, 0.0569249561670159],
		[4e8, 4e8 - 30000, 0.06680450282877265, 0.9331954971712273],
	])('uses the uniform asymptotic expansion for a = %d, x = %d', (a, x, p, q) => {
		const r = incGamma(a, x);
		expectClose(r.lower, p, 1e-13);
		expectClose(r.upper, q, 1e-13);
	});

	it('stays fast and finite for astronomically large shapes', () => {
		const r = incGamma(1e20, 1e20);
		expectClose(r.lower, 0.5 + 1 / (3 * Math.sqrt(2 * Math.PI * 1e20)), 1e-15);
		expect(incGamma(1e20, 1).lower).toBe(0);
		expect(incGamma(1e20, 1e21).upper).toBe(0);
	});

	it('is exact for a = 1 (the exponential distribution)', () => {
		expect(incGamma(1, 2).upper).toBe(Math.exp(-2));
		expect(incGamma(1, 1e-20).lower).toBe(1e-20);
	});

	it('handles the boundaries and invalid arguments', () => {
		expect(incGamma(2, 0)).toEqual({ lower: 0, upper: 1 });
		expect(incGamma(2, Infinity)).toEqual({ lower: 1, upper: 0 });
		expect(incGamma(0, 1).lower).toBeNaN();
		expect(incGamma(1, -1).upper).toBeNaN();
		expect(incGamma(NaN, 1).lower).toBeNaN();
	});

	it('forms the prefactor and density without overflow', () => {
		// x^a e^-x / Gamma(a) at a = x = 1e6 is sqrt(a / 2 pi) e^-stirlerr(a)
		expectClose(gammaPrefix(1e6, 1e6), Math.sqrt(1e6 / (2 * Math.PI)) * Math.exp(-1 / 12e6), 1e-14);
		expectClose(gammaDensity(3, 2), (2 * 2 * Math.exp(-2)) / 2, 1e-15);
		expect(gammaDensity(0.5, 0)).toBe(Infinity);
		expect(gammaDensity(1, 0)).toBe(1);
		expect(gammaDensity(2, 0)).toBe(0);
	});
});

describe('gammaQuantile', () => {
	it.each([
		[0.5, 3, 2.6740603137235603],
		[0.01, 0.5, 0.00015708785790970197 / 2],
		[0.95, 50, 6.217105670200204 / 0.1],
	])('P(a, x) = %d with a = %d', (p, a, x) => {
		expectClose(gammaQuantile(a, p), x, 1e-14);
	});

	it('round-trips both tails over a wide range', () => {
		for (const a of [0.05, 0.5, 1, 2.5, 10, 37.5, 500, 1e5]) {
			for (const p of [1e-10, 1e-4, 0.05, 0.5, 0.95]) {
				const lower = gammaQuantile(a, p);
				// The round trip amplifies the quantile's rounding by x f(x) / P (about 2000 at a = 1e5).
				expectClose(incGamma(a, lower).lower, p, 1e-11);
				const upper = gammaQuantile(a, p, true);
				expectClose(incGamma(a, upper).upper, p, 1e-11);
			}
		}
	});

	it('handles the boundaries', () => {
		expect(gammaQuantile(2, 0)).toBe(0);
		expect(gammaQuantile(2, 1)).toBe(Infinity);
		expect(gammaQuantile(2, 0, true)).toBe(Infinity);
		// The true quantile (about 1e-6000) is below the double range.
		expect(gammaQuantile(0.05, 1e-300)).toBe(0);
	});
});

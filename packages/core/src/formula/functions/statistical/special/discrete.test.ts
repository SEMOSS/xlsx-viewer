import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { bd0, log1pmx, poissonPmf } from './deviance';
import {
	binomCdf,
	binomPmf,
	binomQuantile,
	binomRange,
	exactBinomial,
	hypergeomCdf,
	hypergeomPmf,
	negBinomCdf,
	negBinomPmf,
	poissonCdf,
} from './discrete';

// Exact rational values (Python fractions) rounded to doubles.
const BINOM: [number, number, number, number, number][] = [
	[3, 10, 0.5, 0.1171875, 0.171875],
	[6, 10, 0.5, 0.205078125, 0.828125],
	[0, 20, 0.3, 0.0007979226629761202, 0.0007979226629761202],
	[7, 20, 0.3, 0.16426198521723648, 0.7722717974181605],
	[20, 20, 0.3, 3.4867844009999975e-11, 1],
	[45, 100, 0.47, 0.07380529452032585, 0.38266956737986146],
	[500, 1000, 0.5, 0.0252250181783608, 0.5126125090891804],
	[2, 2000, 0.001, 0.2708059920477052, 0.6766764387615206],
	[1500, 2000, 0.75, 0.02059757143332068, 0.5085832112651493],
];

describe('binomial', () => {
	it.each(BINOM)('k = %d, n = %d, p = %d', (k, n, p, pmf, cdf) => {
		expectClose(binomPmf(k, n, p, 1 - p), pmf, 1e-13);
		expectClose(binomCdf(k, n, p, 1 - p), cdf, 1e-13);
	});

	it('is exact for small n and dyadic p', () => {
		expect(binomPmf(3, 10, 0.5, 0.5)).toBe(0.1171875);
		expect(binomCdf(3, 10, 0.5, 0.5)).toBe(0.171875);
		expect(exactBinomial(56, 28)).toBe(7648690600760440);
		expect(exactBinomial(10, 0)).toBe(1);
	});

	it('handles degenerate probabilities and ranges', () => {
		expect(binomPmf(0, 5, 0, 1)).toBe(1);
		expect(binomPmf(1, 5, 0, 1)).toBe(0);
		expect(binomPmf(5, 5, 1, 0)).toBe(1);
		expect(binomPmf(6, 5, 0.5, 0.5)).toBe(0);
		expect(binomCdf(-1, 5, 0.5, 0.5)).toBe(0);
		expect(binomCdf(5, 5, 0.5, 0.5)).toBe(1);
		expectClose(binomRange(45, 50, 60, 0.75, 0.25), 0.5236297934718872, 1e-13);
		expectClose(
			binomRange(2000, 8000, 10000, 0.5, 0.5),
			1 - binomCdf(1999, 10000, 0.5, 0.5),
			1e-13,
		);
	});

	it('finds the smallest k with P(X <= k) >= alpha', () => {
		expect(binomQuantile(6, 0.5, 0.75)).toBe(4);
		// alpha equal to a cumulative value picks that value's k
		expect(binomQuantile(6, 0.5, 0.65625)).toBe(3);
		expect(binomQuantile(10, 0.5, 0)).toBe(0);
		expect(binomQuantile(10, 0, 0.5)).toBe(0);
		expect(binomQuantile(10, 1, 0.5)).toBe(10);
		expect(binomQuantile(100000, 0.3, 0.5)).toBe(30000);
		const k = binomQuantile(1e6, 0.001, 0.99);
		expect(binomCdf(k, 1e6, 0.001, 0.999)).toBeGreaterThanOrEqual(0.99);
		expect(binomCdf(k - 1, 1e6, 0.001, 0.999)).toBeLessThan(0.99);
	});
});

describe('Poisson', () => {
	it.each([
		[0, 0.5, 0.6065306597126334, 0.6065306597126334],
		[2, 5, 0.08422433748856833, 0.12465201948308115],
		[10, 3.7, 0.003276156388236691, 0.9984278189990823],
		[25, 20, 0.04458764910826788, 0.8878150272820301],
		[100, 100, 0.039860996809147134, 0.5265621985299984],
		[3, 50, 4.0182288499248284e-18, 4.269159205144934e-18],
	])('k = %d, mean = %d', (k, mean, pmf, cdf) => {
		expectClose(poissonPmf(k, mean), pmf, 1e-14);
		expectClose(poissonCdf(k, mean), cdf, 1e-13);
	});

	it('handles a zero mean and large counts', () => {
		expect(poissonPmf(0, 0)).toBe(1);
		expect(poissonPmf(2, 0)).toBe(0);
		expect(poissonCdf(3, 0)).toBe(1);
		expect(poissonCdf(-1, 2)).toBe(0);
		expectClose(poissonCdf(5000, 5000), 0.5037611677708467, 1e-12);
	});
});

describe('hypergeometric and negative binomial', () => {
	it.each([
		[1, 4, 8, 20, 0.3632610939112487, 0.46542827657378744],
		[0, 5, 10, 50, 0.3105627820045687, 0.3105627820045687],
		[3, 10, 30, 100, 0.28116339430253845, 0.6540199886608099],
		[40, 100, 400, 1000, 0.08560998572862107, 0.5450660029793728],
		[5, 5, 5, 60, 1.830994786791643e-7, 1],
	])('hypergeometric x = %d, n = %d, M = %d, N = %d', (x, n, m, total, pmf, cdf) => {
		expectClose(hypergeomPmf(x, n, m, total), pmf, 1e-13);
		expectClose(hypergeomCdf(x, n, m, total), cdf, 1e-13);
	});

	it.each([
		[10, 5, 0.25, 0.05504866037517786, 0.3135140584781766],
		[0, 3, 0.4, 0.06400000000000002, 0.06400000000000002],
		[7, 1, 0.1, 0.04782969, 0.56953279],
		[30, 12, 0.6, 7.929172687691413e-6, 0.9999907966264909],
	])('negative binomial f = %d, s = %d, p = %d', (f, s, p, pmf, cdf) => {
		expectClose(negBinomPmf(f, s, p, 1 - p), pmf, 1e-13);
		expectClose(negBinomCdf(f, s, p, 1 - p), cdf, 1e-13);
	});
});

describe('saddle-point helpers', () => {
	it('computes the deviance term without cancellation', () => {
		expect(bd0(0, 3)).toBe(3);
		expectClose(bd0(10, 10.000001), 4.999999659182688e-14, 1e-9);
		expectClose(bd0(3, 7), 3 * Math.log(3 / 7) + 4, 1e-15);
	});

	it('computes log1p(x) - x accurately', () => {
		expectClose(log1pmx(1e-5), -4.999966666916665e-11, 1e-15);
		expectClose(log1pmx(-0.5), Math.log(0.5) + 0.5, 1e-15);
		expectClose(log1pmx(0.3), Math.log1p(0.3) - 0.3, 1e-14);
		expectClose(log1pmx(5), Math.log(6) - 5, 1e-15);
	});
});

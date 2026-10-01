import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { solveQuantile } from './solve';

// Exponential distribution with rate 1: closed-form quantiles to compare with.
const cdf = (x: number) => -Math.expm1(-x);
const sf = (x: number) => Math.exp(-x);
const pdf = (x: number) => Math.exp(-x);

describe('solveQuantile', () => {
	it('solves distribution functions from good and bad starting points', () => {
		for (const guess of [0.5, 1e-9, 1e9, NaN, -3]) {
			expectClose(
				solveQuantile({ tail: cdf, density: pdf, increasing: true }, 0.25, guess),
				-Math.log1p(-0.25),
				1e-14,
			);
		}
	});

	it('solves survival functions deep in the tail', () => {
		const x = solveQuantile({ tail: sf, density: pdf, increasing: false }, 1e-300, 1);
		expectClose(x, 300 * Math.LN10, 1e-14);
	});

	it('respects a finite upper bound', () => {
		// Uniform distribution on (0, 1): P(X <= x) = x
		const x = solveQuantile(
			{ tail: (v) => v, density: () => 1, increasing: true, upper: 1 },
			0.999,
			5,
		);
		expectClose(x, 0.999, 1e-14);
	});

	it('returns 0 when the root is below the double range', () => {
		// P(X <= x) = x^0.01 has its 1e-300 quantile at 1e-30000.
		const tail = (v: number) => Math.pow(v, 0.01);
		const density = (v: number) => 0.01 * Math.pow(v, -0.99);
		expect(solveQuantile({ tail, density, increasing: true }, 1e-300, 0.5)).toBe(0);
	});
});

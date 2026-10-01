import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { normCdf, normInv, normInvUpper, normPdf, normSf } from './normal';

// [z, P(Z <= z), density] from 60-digit Decimal arithmetic.
const CDF: [number, number, number][] = [
	[-37.5, 4.605353009581955e-308, 1.7282337322841054e-306],
	[-20, 2.7536241186062337e-89, 5.520948362159764e-88],
	[-10, 7.619853024160525e-24, 7.694598626706419e-23],
	[-5, 2.866515718791939e-7, 1.4867195147342977e-6],
	[-3, 0.0013498980316300946, 0.0044318484119380075],
	[-1.96, 0.024997895148220435, 0.05844094433345146],
	[-1, 0.15865525393145705, 0.24197072451914334],
	[-0.5, 0.3085375387259869, 0.35206532676429947],
	[0, 0.5, 0.3989422804014327],
	[0.5, 0.6914624612740131, 0.35206532676429947],
	[1, 0.8413447460685429, 0.24197072451914334],
	[1.96, 0.9750021048517795, 0.05844094433345146],
	[3, 0.9986501019683699, 0.0044318484119380075],
	[5, 0.9999997133484281, 1.4867195147342977e-6],
];

// [p, z] with P(Z <= z) = p.
const QUANTILE: [number, number][] = [
	[1e-300, -37.0470962993612],
	[1e-100, -21.273453560965326],
	[1e-20, -9.262340089798407],
	[1e-10, -6.361340902404057],
	[1e-5, -4.264890793922825],
	[0.001, -3.0902323061678136],
	[0.025, -1.9599639845400543],
	[0.1, -1.2815515655446004],
	[0.3, -0.5244005127080408],
	[0.7, 0.5244005127080407],
	[0.9, 1.2815515655446006],
	[0.975, 1.9599639845400538],
	[0.999, 3.090232306167813],
	[0.9999999, 5.199337582290661],
];

describe('standard normal', () => {
	it.each(CDF)('z = %d', (z, p, f) => {
		expectClose(normCdf(z), p, 1e-14);
		expectClose(normSf(-z), p, 1e-14);
		expectClose(normPdf(z), f, 1e-15);
	});

	it.each(QUANTILE)('quantile at p = %d', (p, z) => {
		expectClose(normInv(p), z, 1e-15);
	});

	it('handles the boundaries', () => {
		expect(normInv(0.5)).toBe(0);
		expect(normInv(0)).toBe(-Infinity);
		expect(normInv(1)).toBe(Infinity);
		expect(normInv(1.5)).toBeNaN();
		expect(normCdf(-40)).toBe(0);
		expect(normCdf(40)).toBe(1);
		expect(normSf(NaN)).toBeNaN();
		expectClose(normInvUpper(1e-50), 14.933337534788489, 1e-15);
	});
});

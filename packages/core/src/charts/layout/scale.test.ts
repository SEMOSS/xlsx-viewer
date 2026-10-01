import { describe, expect, it } from 'vitest';
import { computeValueScale, maxIntervalsFor, niceStep, roundToUnit, scaleFraction } from './scale';

const bounds = (lo: number, hi: number, opts = {}) => {
	const s = computeValueScale(lo, hi, opts);
	return [s.min, s.max, s.majorUnit];
};

describe('niceStep', () => {
	it('rounds up to 1-2-5 steps', () => {
		expect(niceStep(0.33)).toBe(0.5);
		expect(niceStep(9.135)).toBe(10);
		expect(niceStep(11)).toBe(20);
		expect(niceStep(2)).toBe(2);
		expect(niceStep(2.01)).toBe(5);
		expect(niceStep(0.0012)).toBe(0.002);
		expect(niceStep(0)).toBe(1);
	});
});

describe('computeValueScale (Excel automatic axis)', () => {
	it('anchors positive data at zero with 5% headroom', () => {
		expect(bounds(0, 87)).toEqual([0, 100, 10]);
		expect(computeValueScale(0, 87).ticks).toEqual([0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]);
		expect(bounds(45, 71)).toEqual([0, 80, 10]);
		expect(bounds(0, 100)).toEqual([0, 120, 20]);
		expect(bounds(1.2, 3.2)).toEqual([0, 3.5, 0.5]);
		expect(bounds(0, 1000)).toEqual([0, 1200, 200]);
	});

	it('starts above zero when the data is a narrow band near the top', () => {
		expect(bounds(90, 100)).toEqual([84, 102, 2]);
		expect(bounds(1000, 1100)).toEqual([940, 1120, 20]);
	});

	it('handles negative and mixed ranges', () => {
		expect(bounds(-20, 80)).toEqual([-40, 100, 20]);
		expect(bounds(-87, -10)).toEqual([-100, 0, 10]);
		expect(bounds(-100, -90)).toEqual([-102, -84, 2]);
		expect(bounds(-5, 10)).toEqual([-6, 12, 2]);
	});

	it('scales 100% stacks to exactly 0..100%', () => {
		expect(bounds(0, 1, { percent: true })).toEqual([0, 1, 0.1]);
		expect(bounds(-0.4, 1, { percent: true })).toEqual([-0.4, 1, 0.2]);
	});

	it('honours fixed bounds and units', () => {
		expect(bounds(0, 87, { majorUnit: 25 })).toEqual([0, 100, 25]);
		expect(bounds(0, 87, { min: 10, max: 50 })).toEqual([10, 50, 5]);
		expect(bounds(0, 87, { max: 60 })).toEqual([0, 60, 10]);
		expect(bounds(20, 87, { min: 50 })).toEqual([50, 95, 5]);
		const minor = computeValueScale(0, 87, { minorUnit: 5 });
		expect(minor.minorTicks.slice(0, 3)).toEqual([5, 15, 25]);
	});

	it('uses fewer intervals on short axes', () => {
		expect(bounds(0, 87, { maxIntervals: 5 })).toEqual([0, 100, 20]);
		expect(bounds(0, 3.3, { maxIntervals: 1 })).toEqual([0, 5, 5]);
		expect(bounds(0, 3.3, { maxIntervals: 4 })).toEqual([0, 4, 1]);
	});

	it('handles flat and empty data', () => {
		expect(bounds(5, 5)).toEqual([0, 6, 1]);
		expect(bounds(-5, -5)).toEqual([-6, 0, 1]);
		expect(bounds(0, 0)).toEqual([0, 1, 0.1]);
		expect(bounds(NaN, NaN)).toEqual([0, 1, 0.1]);
	});

	it('produces clean tick values', () => {
		const s = computeValueScale(0, 0.28);
		expect(s.ticks).toEqual([0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3]);
		expect(roundToUnit(0.1 + 0.2, 0.1)).toBe(0.3);
	});

	it('computes logarithmic scales', () => {
		const s = computeValueScale(3, 870, { logBase: 10 });
		expect([s.min, s.max, s.logBase]).toEqual([1, 1000, 10]);
		expect(s.ticks).toEqual([1, 10, 100, 1000]);
		expect(scaleFraction(s, 100)).toBeCloseTo(2 / 3);
		expect(computeValueScale(1, 30, { logBase: 2 }).ticks).toEqual([1, 2, 4, 8, 16, 32]);
	});
});

describe('maxIntervalsFor', () => {
	it('fits labels along the axis, at most 10 intervals', () => {
		expect(maxIntervalsFor(200, 20)).toBe(10);
		expect(maxIntervalsFor(1000, 20)).toBe(10);
		expect(maxIntervalsFor(83, 20)).toBe(4);
		expect(maxIntervalsFor(10, 20)).toBe(1);
		expect(maxIntervalsFor(0, 20)).toBe(10);
	});
});

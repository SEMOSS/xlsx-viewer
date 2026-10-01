import { describe, expect, it } from 'vitest';
import { dateToSerial } from '../../values/dates';
import { computeDateScale, dateFraction, detectBaseUnit } from './date-scale';

const d = (y: number, m: number, day: number) => dateToSerial(y, m, day);

describe('date axes', () => {
	it('detects the base unit from the dates', () => {
		expect(detectBaseUnit([d(2024, 1, 1), d(2024, 1, 2), d(2024, 1, 5)])).toBe('days');
		expect(detectBaseUnit([d(2024, 1, 1), d(2024, 2, 1), d(2024, 4, 1)])).toBe('months');
		expect(detectBaseUnit([d(2020, 1, 1), d(2021, 1, 1)])).toBe('years');
	});

	it('places ticks on whole units and positions dates proportionally', () => {
		const months = [d(2024, 1, 1), d(2024, 2, 1), d(2024, 3, 1), d(2024, 4, 1)];
		const scale = computeDateScale(months, { between: false });
		expect(scale.baseUnit).toBe('months');
		expect(scale.ticks).toEqual(months);
		expect(dateFraction(scale, months[0])).toBe(0);
		expect(dateFraction(scale, months[3])).toBe(1);
		const bars = computeDateScale(months, { between: true });
		expect(bars.min).toBeLessThan(months[0]);
		expect(bars.slotDays).toBeCloseTo(30.44, 1);
		const days = Array.from({ length: 60 }, (_, i) => d(2024, 1, 1) + i);
		expect(computeDateScale(days, { between: false, maxIntervals: 10 }).ticks).toHaveLength(9);
	});
});

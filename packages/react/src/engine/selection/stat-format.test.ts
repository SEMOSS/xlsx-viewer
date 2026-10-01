import { describe, expect, it } from 'vitest';

import { formatStatistic } from './stat-format';

describe('formatStatistic', () => {
	it('writes General numbers in the conventions of the UI locale', () => {
		expect(formatStatistic(1234.5, 'General', 'en')).toBe('1,234.5');
		expect(formatStatistic(1234.5, 'General', 'de')).toBe('1.234,5');
		expect(formatStatistic(1234.5, 'General', 'fr')).toBe('1\u202f234,5');
		expect(formatStatistic(12345.5, 'General', 'es')).toBe('12.345,5');
		expect(formatStatistic(1234.5, 'General', 'zh-CN')).toBe('1,234.5');
	});

	it('keeps up to ten decimals and switches to scientific notation for huge values', () => {
		expect(formatStatistic(1 / 3, 'general', 'en')).toBe('0.3333333333');
		expect(formatStatistic(1.5e20, 'General', 'en')).toBe('1.5E20');
	});

	it("applies the active cell's explicit number format like the cell shows it", () => {
		expect(formatStatistic(1234.5, '#,##0.00', 'de')).toBe('1,234.50');
		expect(formatStatistic(0.25, '0%', 'fr')).toBe('25%');
		expect(formatStatistic(45292, 'yyyy-mm-dd', 'zh-CN')).toBe('2024-01-01');
	});
});

import { describe, expect, it } from 'vitest';

import { formatNumber, formatPercent, numberFormatter } from './number-format';

describe('number-format', () => {
	it('formats numbers with the separators of each locale', () => {
		expect(formatNumber(1234567.25, 'en')).toBe('1,234,567.25');
		expect(formatNumber(1234567.25, 'de')).toBe('1.234.567,25');
		expect(formatNumber(1234567.25, 'fr')).toBe('1\u202f234\u202f567,25');
		expect(formatNumber(1234567.25, 'es')).toBe('1.234.567,25');
		expect(formatNumber(1234567.25, 'zh-CN')).toBe('1,234,567.25');
	});

	it('writes whole percentages the way each locale does', () => {
		expect(formatPercent(125, 'en')).toBe('125%');
		expect(formatPercent(100, 'de')).toBe('100\u00a0%');
		expect(formatPercent(75, 'fr')).toBe('75\u00a0%');
		expect(formatPercent(50, 'es')).toBe('50\u00a0%');
		expect(formatPercent(200, 'zh-CN')).toBe('200%');
	});

	it('reuses formatters for the same locale and options', () => {
		expect(numberFormatter('de', { maximumFractionDigits: 1 })).toBe(
			numberFormatter('de', { maximumFractionDigits: 1 }),
		);
		expect(numberFormatter('de')).not.toBe(numberFormatter('fr'));
	});
});

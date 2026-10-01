import { describe, expect, it } from 'vitest';

import { dateToSerial, parseNumericText, serialToDateParts } from './index';

describe('date serials (1900 system)', () => {
	it.each([
		[1900, 1, 1, 1],
		[1900, 2, 28, 59],
		[1900, 3, 1, 61],
		[2000, 1, 1, 36526],
		[2024, 2, 29, 45351],
		[9999, 12, 31, 2958465],
	])('%i-%i-%i is serial %i', (y, m, d, serial) => {
		expect(dateToSerial(y, m, d)).toBe(serial);
		const p = serialToDateParts(serial);
		expect([p.year, p.month, p.day]).toEqual([y, m, d]);
	});

	it('reports the phantom 1900-02-29 and serial 0', () => {
		expect(serialToDateParts(60)).toMatchObject({ year: 1900, month: 2, day: 29 });
		expect(serialToDateParts(0)).toMatchObject({ year: 1900, month: 1, day: 0 });
	});

	it('matches Excel WEEKDAY numbering around the leap-year bug', () => {
		expect(serialToDateParts(1).weekday).toBe(0); // Excel: Sunday
		expect(serialToDateParts(61).weekday).toBe(4); // 1900-03-01 Thursday
		expect(serialToDateParts(45351).weekday).toBe(4); // 2024-02-29 Thursday
	});

	it('splits time of day', () => {
		const p = serialToDateParts(45351.75);
		expect([p.hour, p.minute, p.second]).toEqual([18, 0, 0]);
	});

	it('supports the 1904 system', () => {
		expect(dateToSerial(1904, 1, 1, true)).toBe(0);
		expect(serialToDateParts(1, true)).toMatchObject({ year: 1904, month: 1, day: 2 });
	});
});

describe('parseNumericText', () => {
	it.each([
		['42', 42, undefined],
		['-1.5e3', -1500, '0.00E+00'],
		['1,234', 1234, '#,##0'],
		['1,234.50', 1234.5, '#,##0.00'],
		['50%', 0.5, '0%'],
		['12.5%', 0.125, '0.00%'],
		['$1,234.56', 1234.56, '"$"#,##0.00'],
		['-$5', -5, '"$"#,##0'],
		['(100)', -100, undefined],
		['1 1/2', 1.5, '# ?/?'],
	])('%s', (text, value, numFmt) => {
		const r = parseNumericText(text);
		expect(r?.value).toBeCloseTo(value as number, 10);
		expect(r?.numFmt).toBe(numFmt);
	});

	it('parses dates and times', () => {
		expect(parseNumericText('1/2/2024')).toMatchObject({ value: 45293, numFmt: 'm/d/yyyy' });
		expect(parseNumericText('2024-01-02')?.value).toBe(45293);
		expect(parseNumericText('2-Jan-2024')).toMatchObject({ value: 45293, numFmt: 'd-mmm-yy' });
		expect(parseNumericText('Jan 2, 2024')?.value).toBe(45293);
		expect(parseNumericText('13:30')).toMatchObject({ value: 13.5 / 24, numFmt: 'h:mm' });
		expect(parseNumericText('1:30 PM')).toMatchObject({ value: 13.5 / 24, numFmt: 'h:mm AM/PM' });
		expect(parseNumericText('1/2/2024 6:00')).toMatchObject({
			value: 45293.25,
			numFmt: 'm/d/yyyy h:mm',
		});
		expect(parseNumericText('2/30/2024')).toBeNull();
	});

	it('rejects non-numeric text', () => {
		for (const t of ['abc', '1,2', '12abc', '$', '%', '--1', 'TRUE', '']) {
			expect(parseNumericText(t)).toBeNull();
		}
	});
});

import { describe, expect, it } from 'vitest';

import { formatValue } from './index';

const text = (value: number, code: string, date1904 = false) =>
	formatValue(value, code, { date1904 }).text;

/** Serial of a time of day. */
const time = (h: number, m: number, s: number) => (h * 3600 + m * 60 + s) / 86_400;

/** 2023-03-15 (a Wednesday) and 2023-03-05 (a Sunday). */
const MAR_15 = 45000;
const MAR_5 = 44990;

describe('date codes', () => {
	it.each([
		['yyyy-mm-dd', MAR_15, '2023-03-15'],
		['m/d/yyyy', MAR_15, '3/15/2023'],
		['yy', MAR_15, '23'],
		['y', MAR_15, '23'],
		['yyy', MAR_15, '2023'],
		['yyyyy', MAR_15, '2023'],
		['m', MAR_15, '3'],
		['mm', MAR_15, '03'],
		['mmm', MAR_15, 'Mar'],
		['mmmm', MAR_15, 'March'],
		['mmmmm', MAR_15, 'M'],
		['d', MAR_5, '5'],
		['dd', MAR_5, '05'],
		['ddd', MAR_5, 'Sun'],
		['dddd', MAR_5, 'Sunday'],
		['dddd', MAR_15, 'Wednesday'],
		['mmm d, yyyy', MAR_5, 'Mar 5, 2023'],
		['dddd, mmmm dd, yyyy', MAR_5, 'Sunday, March 05, 2023'],
		['d-mmm-yy', MAR_15, '15-Mar-23'],
		['YYYY/MM/DD', MAR_15, '2023/03/15'],
		['e', MAR_15, '2023'],
		['bbbb', MAR_15, '2566'],
		['bb', MAR_15, '66'],
		['yyyy.mm.dd', MAR_15, '2023.03.15'],
		['m/d/yyyy', 1, '1/1/1900'],
		['m/d/yyyy', 0, '1/0/1900'],
		['m/d/yyyy', 59, '2/28/1900'],
		['m/d/yyyy dddd', 60, '2/29/1900 Wednesday'],
		['m/d/yyyy', 61, '3/1/1900'],
		['m/d/yyyy dddd', 36526, '1/1/2000 Saturday'],
		['m/d/yyyy dddd', 45351, '2/29/2024 Thursday'],
		['yyyy-mm-dd dddd', 2958465, '9999-12-31 Friday'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});
});

describe('time codes', () => {
	it.each([
		['h:mm', time(10, 29, 40), '10:29'],
		['h:mm:ss', time(10, 29, 40), '10:29:40'],
		['hh:mm AM/PM', time(10, 29, 40), '10:29 AM'],
		['h:m:s', time(13, 5, 9), '13:5:9'],
		['hh:mm:ss', time(13, 5, 9), '13:05:09'],
		['h:mm AM/PM', 0, '12:00 AM'],
		['h:mm AM/PM', 0.5, '12:00 PM'],
		['h:mm AM/PM', time(13, 0, 0), '1:00 PM'],
		['h:mm AM/PM', time(1, 0, 0), '1:00 AM'],
		['hh:mm AM/PM', 0.7, '04:48 PM'],
		['h:mm am/pm', 0.75, '6:00 pm'],
		['h:mm A/P', 0.25, '6:00 A'],
		['h a/p', 0.75, '6 p'],
		['h "h" m "min"', time(2, 5, 0), '2 h 5 min'],
		['mm:ss', time(0, 1, 0), '01:00'],
		['m:ss', time(0, 1, 5), '1:05'],
		['mmss.0', time(13, 30, 55.2), '3055.2'],
		['mm:ss.00', 1.234 / 86_400, '00:01.23'],
		['ss.000', 1.5 / 86_400, '01.500'],
		['h:mm:ss.000', time(13, 30, 55.123), '13:30:55.123'],
		['hh:mm:ss.00', 0.3456, '08:17:39.84'],
		['h:mm:ss.0000', time(1, 2, 3.4567), '1:02:03.4570'],
		['h:mm', MAR_15 + 0.75, '18:00'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});
});

describe('elapsed time', () => {
	it.each([
		['[h]:mm', 1.25, '30:00'],
		['[h]:mm:ss', 2.5, '60:00:00'],
		['[h]:mm:ss', 2.3456, '56:17:40'],
		['[h]', 0.5, '12'],
		['[hh]', time(1, 0, 0), '01'],
		['[mm]:ss', time(1, 30, 0), '90:00'],
		['[m]', 1, '1440'],
		['[ss]', 0.001, '86'],
		['[ss].0', 0.001, '86.4'],
		['[s]', time(0, 2, 3), '123'],
		['[h]:mm:ss', 0, '0:00:00'],
		['[h]:[mm]', 1.5 + time(0, 7, 0), '36:07'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});
});

describe('rounding to the displayed precision', () => {
	it('rounds to the second, then truncates larger units', () => {
		expect(text(time(23, 59, 59.6), 'h:mm:ss')).toBe('0:00:00');
		expect(text(MAR_15 + time(23, 59, 59.6), 'm/d/yyyy h:mm:ss')).toBe('3/16/2023 0:00:00');
		expect(text(0.99999999, 'h:mm')).toBe('0:00');
		expect(text(time(23, 59, 30), 'h:mm')).toBe('23:59');
		expect(text(time(10, 29, 59.6), 'h:mm')).toBe('10:30');
		expect(text(time(0, 0, 0.44), 'h:mm:ss')).toBe('0:00:00');
		expect(text(time(0, 0, 0.44), 'h:mm:ss.0')).toBe('0:00:00.4');
		expect(text(time(0, 0, 59.96), 'mm:ss.0')).toBe('01:00.0');
	});
});

describe('locale and system tags', () => {
	it.each([
		['[$-409]mmmm d, yyyy;@', MAR_15, 'March 15, 2023'],
		['[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy', MAR_5, 'Sunday, March 5, 2023'],
		['[$-x-sysdate]dddd, mmmm dd, yyyy', MAR_15, 'Wednesday, March 15, 2023'],
		['[$-F400]h:mm:ss\\ AM/PM', 0.5, '12:00:00 PM'],
		['[$-x-systime]h:mm:ss AM/PM', time(13, 30, 55), '1:30:55 PM'],
		['[$-404]e/m/d', MAR_15, '2023/3/15'],
		['[$-411]ggge"年"m"月"d"日"', MAR_15, '2023年3月15日'],
		['[$-409]h:mm AM/PM;@', time(9, 5, 0), '9:05 AM'],
		['[Red]m/d/yyyy', MAR_15, '3/15/2023'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});

	it('keeps the section color', () => {
		expect(formatValue(MAR_15, '[Blue]m/d/yyyy').color).toBe('#0000FF');
	});
});

describe('1904 date system', () => {
	it('shifts the epoch', () => {
		expect(text(0, 'm/d/yyyy', true)).toBe('1/1/1904');
		expect(text(1, 'm/d/yyyy', true)).toBe('1/2/1904');
		expect(text(MAR_15, 'm/d/yyyy', true)).toBe('3/16/2027');
		expect(text(2957003, 'yyyy-mm-dd', true)).toBe('9999-12-31');
		expect(formatValue(2957004, 'yyyy-mm-dd', { date1904: true }).invalid).toBe(true);
	});

	it('shows negative date-times with a minus sign', () => {
		expect(text(-0.5, '[h]:mm', true)).toBe('-12:00');
		expect(text(-1.25, 'm/d/yyyy h:mm', true)).toBe('-1/2/1904 6:00');
		expect(text(-time(1, 30, 0), 'h:mm', true)).toBe('-1:30');
	});
});

describe('values a date format cannot show', () => {
	it.each([
		['m/d/yyyy', -1],
		['h:mm', -0.5],
		['[h]:mm', -0.1],
		['yyyy-mm-dd', 2958466],
		['h:mm:ss', 3e6],
		['m/d/yyyy', 1e308],
	])('%s with %s is #####', (code, value) => {
		expect(formatValue(value, code)).toEqual({ text: '#####', invalid: true });
	});

	it('formats the absolute value in a negative section', () => {
		expect(text(-1, 'm/d/yyyy;m/d/yyyy')).toBe('1/1/1900');
	});

	it('passes text through or into the text section', () => {
		expect(formatValue('abc', 'm/d/yyyy').text).toBe('abc');
		expect(formatValue('abc', 'm/d/yyyy;@').text).toBe('abc');
		expect(formatValue('abc', 'm/d/yyyy;;;"<"@">"').text).toBe('<abc>');
	});
});

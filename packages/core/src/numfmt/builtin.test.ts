import { describe, expect, it } from 'vitest';

import { BUILTIN_NUMBER_FORMATS } from '../styles/builtin-number-formats';
import { formatValue } from './index';
import type { FormattedValue } from './index';

/** 2023-03-15 13:30:55 */
const STAMP = 45000 + (13 * 3600 + 30 * 60 + 55) / 86_400;

const fmt = (id: number, value: number | string) => formatValue(value, BUILTIN_NUMBER_FORMATS[id]);

describe('built-in number formats (en-US)', () => {
	it.each<[number, number | string, string]>([
		[0, 1234.5, '1234.5'],
		[0, -0.25, '-0.25'],
		[1, 1234.5, '1235'],
		[1, -1234.5, '-1235'],
		[2, 1234.5, '1234.50'],
		[2, 0.005, '0.01'],
		[3, 1234.5, '1,235'],
		[3, -1234567, '-1,234,567'],
		[4, 1234.5, '1,234.50'],
		[4, -0.5, '-0.50'],
		[5, 1234.5, '$1,235 '],
		[5, -1234.5, '($1,235)'],
		[5, 0, '$0 '],
		[6, 1234.5, '$1,235 '],
		[6, -1234.5, '($1,235)'],
		[7, 1234.567, '$1,234.57 '],
		[7, -1234.567, '($1,234.57)'],
		[8, -1234.5, '($1,234.50)'],
		[8, 0, '$0.00 '],
		[9, 0.1234, '12%'],
		[9, -0.5, '-50%'],
		[10, 0.1234, '12.34%'],
		[10, 1, '100.00%'],
		[11, 1234.5, '1.23E+03'],
		[11, 0.000123, '1.23E-04'],
		[12, 1.25, '1 1/4'],
		[12, 5, '5    '],
		[13, 1.25, '1  1/4 '],
		[13, 0.123, '  8/65'],
		[14, STAMP, '3/15/2023'],
		[15, STAMP, '15-Mar-23'],
		[16, STAMP, '15-Mar'],
		[17, STAMP, 'Mar-23'],
		[18, STAMP, '1:30 PM'],
		[19, STAMP, '1:30:55 PM'],
		[20, STAMP, '13:30'],
		[21, STAMP, '13:30:55'],
		[22, STAMP, '3/15/2023 13:30'],
		[27, STAMP, '2023/3/15'],
		[30, STAMP, '3/15/23'],
		[36, STAMP, '2023/3/15'],
		[37, 1234, '1,234 '],
		[37, -1234, '(1,234)'],
		[38, -1234, '(1,234)'],
		[39, 1234.5, '1,234.50 '],
		[39, -1234.5, '(1,234.50)'],
		[40, -1234.5, '(1,234.50)'],
		[45, STAMP, '30:55'],
		[46, 1 + (13 * 3600 + 30 * 60 + 55) / 86_400, '37:30:55'],
		[47, (13 * 3600 + 30 * 60 + 55.2) / 86_400, '3055.2'],
		[48, 12345, '12.3E+3'],
		[48, 0.0001, '100.0E-6'],
		[49, 1234.5, '1234.5'],
		[49, 'abc', 'abc'],
		[50, STAMP, '2023/3/15'],
		[57, STAMP, '2023/3/15'],
		[59, 5, '5'],
		[60, 5, '5.00'],
		[61, 1234, '1,234'],
		[62, 1234, '1,234.00'],
		[67, 0.5, '50%'],
		[68, 0.5, '50.00%'],
		[69, 1.5, '1 1/2'],
		[70, 1.5, '1  1/2 '],
	])('id %i with %s -> "%s"', (id, value, expected) => {
		expect(fmt(id, value).text).toBe(expected);
	});

	it('colors the negative section of ids 6, 8, 38 and 40 only', () => {
		for (const id of [6, 8, 38, 40]) {
			expect(fmt(id, -1).color).toBe('#FF0000');
			expect(fmt(id, 1).color).toBeUndefined();
		}
	});

	it('formats every built-in code without throwing', () => {
		for (const code of Object.values(BUILTIN_NUMBER_FORMATS)) {
			for (const v of [0, 1.5, -1.5, 45000.25, 'x']) {
				expect(typeof formatValue(v, code).text).toBe('string');
			}
		}
	});
});

describe('accounting formats 41-44: fill and alignment spaces', () => {
	const fill = (before: string, after: string): FormattedValue => ({
		text: before + after,
		fill: { before, char: ' ', after },
	});

	it.each<[number, number, FormattedValue]>([
		[41, 1234, fill(' ', '1,234 ')],
		[41, -1234, fill(' ', '(1,234)')],
		[41, 0, fill(' ', '- ')],
		[41, 0.4, fill(' ', '0 ')],
		[42, 1234, fill(' $', '1,234 ')],
		[42, -1234, fill(' $', '(1,234)')],
		[42, 0, fill(' $', '- ')],
		[43, 1234.5, fill(' ', '1,234.50 ')],
		[43, -1234.5, fill(' ', '(1,234.50)')],
		[43, 0, fill(' ', '-   ')],
		[44, 1234.5, fill(' $', '1,234.50 ')],
		[44, -1234.5, fill(' $', '(1,234.50)')],
		[44, 0, fill(' $', '-   ')],
		[44, 1234567.891, fill(' $', '1,234,567.89 ')],
	])('id %i with %s', (id, value, expected) => {
		expect(fmt(id, value)).toEqual(expected);
	});

	it('pads text with the text section', () => {
		for (const id of [41, 42, 43, 44]) {
			expect(fmt(id, 'abc')).toEqual({ text: ' abc ' });
		}
	});
});

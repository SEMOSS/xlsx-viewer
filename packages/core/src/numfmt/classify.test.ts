import { describe, expect, it } from 'vitest';

import { BUILTIN_NUMBER_FORMATS } from '../styles/builtin-number-formats';
import { classifyFormat, getFormatDecimals, isDateFormat, isTimeOnlyFormat } from './index';
import type { FormatCategory } from './index';

describe('isDateFormat', () => {
	it.each([
		'm/d/yyyy',
		'[h]:mm:ss',
		'mm:ss',
		'[mm]:ss',
		'[ss]',
		'yyyy',
		'dddd',
		'h AM/PM',
		'e',
		'[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy',
		'[$-409]mmmm d, yyyy;@',
		'[Red]m/d/yyyy',
		'"Due: "m/d/yyyy',
		'[$-404]e/m/d',
	])('%s is a date format', (code) => {
		expect(isDateFormat(code)).toBe(true);
	});

	it.each([
		'General',
		'',
		'0.00',
		'"Date: "0',
		'"yyyy"',
		'[Red]0',
		'[$-409]0.00',
		'\\d0',
		'\\h\\h0',
		'0.00E+00',
		'# ?/?',
		'@',
		'_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
		'0 "days"',
		'[$€-407] #,##0.00',
		'0 g',
	])('%s is not a date format', (code) => {
		expect(isDateFormat(code)).toBe(false);
	});
});

describe('isTimeOnlyFormat', () => {
	it.each<[string, boolean]>([
		['h:mm', true],
		['[h]:mm:ss', true],
		['mm:ss.0', true],
		['h:mm AM/PM', true],
		['[$-F400]h:mm:ss\\ AM/PM', true],
		['m/d/yyyy h:mm', false],
		['m/d/yyyy', false],
		['mmm', false],
		['0.00', false],
	])('%s -> %s', (code, expected) => {
		expect(isTimeOnlyFormat(code)).toBe(expected);
	});
});

describe('classifyFormat', () => {
	it.each<[number, FormatCategory]>([
		[0, 'general'],
		[1, 'number'],
		[2, 'number'],
		[3, 'number'],
		[4, 'number'],
		[5, 'currency'],
		[6, 'currency'],
		[7, 'currency'],
		[8, 'currency'],
		[9, 'percentage'],
		[10, 'percentage'],
		[11, 'scientific'],
		[12, 'fraction'],
		[13, 'fraction'],
		[14, 'date'],
		[15, 'date'],
		[16, 'date'],
		[17, 'date'],
		[18, 'time'],
		[19, 'time'],
		[20, 'time'],
		[21, 'time'],
		[22, 'date'],
		[27, 'date'],
		[30, 'date'],
		[37, 'number'],
		[38, 'number'],
		[39, 'number'],
		[40, 'number'],
		[41, 'accounting'],
		[42, 'accounting'],
		[43, 'accounting'],
		[44, 'accounting'],
		[45, 'time'],
		[46, 'time'],
		[47, 'time'],
		[48, 'scientific'],
		[49, 'text'],
	])('built-in %i is %s', (id, category) => {
		expect(classifyFormat(BUILTIN_NUMBER_FORMATS[id])).toBe(category);
	});

	it.each<[string, FormatCategory]>([
		['', 'general'],
		['GENERAL', 'general'],
		['00000', 'special'],
		['00000-0000', 'special'],
		['[<=9999999]###-####;(###) ###-####', 'special'],
		['000-00-0000', 'special'],
		['0.00;[Red]0.00', 'number'],
		['0.00_ ;[Red]\\-0.00\\ ', 'number'],
		['#,##0.00;-#,##0.00', 'number'],
		['"$"#,##0.00', 'currency'],
		['[$€-407] #,##0.00', 'currency'],
		['#,##0.00 [$€-1]', 'currency'],
		['#,##0.00 "€"', 'currency'],
		['[$£-809]#,##0.00;[Red]\\-[$£-809]#,##0.00', 'currency'],
		['_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-', 'accounting'],
		['0.000%', 'percentage'],
		['0E+00', 'scientific'],
		['# ?/8', 'fraction'],
		['# ??/100', 'fraction'],
		['[$-409]mmmm d, yyyy;@', 'date'],
		['[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy', 'date'],
		['[$-409]h:mm AM/PM;@', 'time'],
		['mm:ss.0;@', 'time'],
		['0 "kg"', 'custom'],
		['#,##0,', 'custom'],
		['0.0,,"M"', 'custom'],
		['[>100]0', 'custom'],
		['[Red]0', 'custom'],
		['yyyy"年"m"月"', 'custom'],
		['"Due: "m/d/yyyy', 'custom'],
		['"Name: "@', 'custom'],
		[';;;', 'custom'],
		['[Red]General', 'custom'],
		['0%%', 'custom'],
	])('%s is %s', (code, category) => {
		expect(classifyFormat(code)).toBe(category);
	});
});

describe('getFormatDecimals', () => {
	it.each<[string, number]>([
		['0', 0],
		['0.00', 2],
		['#,##0.000', 3],
		['0.0%', 1],
		['0.00E+00', 2],
		['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 2],
		['General', 0],
		['@', 0],
		['h:mm:ss.000', 3],
		['m/d/yyyy', 0],
		['# ?/?', 0],
		['#.##', 2],
	])('%s -> %i', (code, decimals) => {
		expect(getFormatDecimals(code)).toBe(decimals);
	});
});

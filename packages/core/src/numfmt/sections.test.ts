import { describe, expect, it } from 'vitest';

import { ERRORS } from '../values/errors';
import { formatValue } from './index';

const text = (value: number | string | boolean, code: string) => formatValue(value, code).text;

describe('positive / negative / zero / text sections', () => {
	it.each([
		['0.00;(0.00)', 5, '5.00'],
		['0.00;(0.00)', -5, '(5.00)'],
		['0.00;(0.00)', 0, '0.00'],
		['0;0', -1.1, '1'],
		['0;-0;"zero"', 0, 'zero'],
		['0;-0;"zero"', -3, '-3'],
		['0;-0;"zero"', 3, '3'],
		['0.00;(0.00);"zero"', -0.001, '(0.00)'],
		['0.00;(0.00);"zero"', 0.001, '0.00'],
		['"foo";"bar";"baz";"qux"', 1, 'foo'],
		['"foo";"bar";"baz";"qux"', -1, 'bar'],
		['"foo";"bar";"baz";"qux"', 0, 'baz'],
		['"foo";"bar";"baz";"qux"', 'sheetjs', 'qux'],
		['"foo";"bar";"baz"', 'sheetjs', 'sheetjs'],
		['"foo";"bar";@', 0, 'foo'],
		['"foo";"bar";@', -1, 'bar'],
		['"foo";"bar";@', 'sheetjs', 'sheetjs'],
		['0;;', -5, ''],
		['0;;', 0, ''],
		['0;;', 5, '5'],
		[';;;', 5, ''],
		[';;;', 'x', ''],
		[';;;', true, ''],
		['0;-0;;@', 0, ''],
		['0;-0;0;', 'x', ''],
		['General;-General', -5, '-5'],
		['General;[Red]General', -5, '5'],
		['#,##0.00;-#,##0.00;"-"', 0, '-'],
		['0;0;0;"T"', 'x', 'T'],
		['@', 1234.5, '1234.5'],
		['@', -1234.5, '-1234.5'],
		['@@', 'ab', 'abab'],
		['@@', 1, '1'],
		['"Name: "@', 'Ann', 'Name: Ann'],
		['"Name: "@', 5, '5'],
		['0;0;0;@" x"', 'abc', 'abc x'],
		['0.00', 'text', 'text'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});
});

describe('conditions', () => {
	it.each([
		['[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', 1234567, '1.2M'],
		['[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', 12345, '12.3K'],
		['[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', 500, '500'],
		['[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', -5000, '-5000'],
		['[<=9999999]###-####;(###) ###-####', 8675309, '867-5309'],
		['[<=9999999]###-####;(###) ###-####', 2813308004, '(281) 330-8004'],
		['[<0]"neg";[>0]"pos";"zero"', -1, 'neg'],
		['[<0]"neg";[>0]"pos";"zero"', 1, 'pos'],
		['[<0]"neg";[>0]"pos";"zero"', 0, 'zero'],
		['[=1]"one";"other"', 2, 'other'],
		['[<>0]0;"zero"', 0, 'zero'],
		['[<>0]0;"zero"', 7, '7'],
		['[>100]0.00', 500, '500.00'],
		['[>100]0.00', 5, '5'],
		['[>=100]0;[<0](0);0.00', -5, '(5)'],
		['[>=100]0;[<0](0);0.00', 5, '5.00'],
		['[<-1.5]0.0;0', -2, '2.0'],
		['[>=1e3]"big";0', 2000, 'big'],
		['[Red][<=100];[Blue][>100]', 50, '50'],
		['[Red][<=100];[Blue][>100]', -5, '-5'],
		['[Red][<=100];[Blue][>100]', 500, '500'],
	])('%s with %s -> "%s"', (code, value, expected) => {
		expect(text(value, code)).toBe(expected);
	});

	it('leaves the sign to sections that only receive negative numbers', () => {
		const code = '[Red][<=-25]General;[Blue][>=25]General;[Green]General;[Yellow]General';
		expect(formatValue(-26, code)).toEqual({ text: '26', color: '#FF0000' });
		expect(formatValue(30, code)).toEqual({ text: '30', color: '#0000FF' });
		expect(formatValue(-1, code)).toEqual({ text: '-1', color: '#00FF00' });
		expect(formatValue('foo', code)).toEqual({ text: 'foo', color: '#FFFF00' });
	});

	it('shows ##### when two conditional sections both miss', () => {
		expect(formatValue(3, '[=1]"one";[=2]"two"')).toEqual({ text: '#####', invalid: true });
		expect(text(2, '[=1]"one";[=2]"two"')).toBe('two');
	});

	it('applies the section colors of tag-only sections', () => {
		expect(formatValue(50, '[Red][<=100];[Blue][>100]').color).toBe('#FF0000');
		expect(formatValue(500, '[Red][<=100];[Blue][>100]').color).toBe('#0000FF');
	});
});

describe('colors', () => {
	it.each([
		['[Black]0', '#000000'],
		['[Blue]0', '#0000FF'],
		['[Cyan]0', '#00FFFF'],
		['[Green]0', '#00FF00'],
		['[Magenta]0', '#FF00FF'],
		['[Red]0', '#FF0000'],
		['[White]0', '#FFFFFF'],
		['[Yellow]0', '#FFFF00'],
		['[RED]0', '#FF0000'],
		['[Color1]0', '#000000'],
		['[Color3]0', '#FF0000'],
		['[Color10]0', '#008000'],
		['[color 16]0', '#808080'],
		['[Color56]0', '#333333'],
		['0[Red]', '#FF0000'],
	])('%s -> %s', (code, color) => {
		expect(formatValue(5, code).color).toBe(color);
	});

	it('ignores unknown colors and uses only the first tag', () => {
		expect(formatValue(5, '[Color57]0').color).toBeUndefined();
		expect(formatValue(5, '[Orange]0')).toEqual({ text: '5' });
		expect(formatValue(5, '[Blue][Red]0').color).toBe('#0000FF');
	});

	it('colors only the section that formats the value', () => {
		expect(formatValue(5, '0;[Red]0')).toEqual({ text: '5' });
		expect(formatValue(-5, '0;[Red]0')).toEqual({ text: '5', color: '#FF0000' });
		expect(formatValue('abc', '[Red]0').color).toBeUndefined();
		expect(formatValue('abc', '[Blue]@').color).toBe('#0000FF');
	});

	it('resolves [ColorN] against a workbook palette', () => {
		const palette = Array.from({ length: 64 }, () => 'FF000000');
		palette[17] = 'FF123456';
		expect(formatValue(5, '[Color10]0', { indexedColors: palette }).color).toBe('#123456');
	});
});

describe('non-numeric values', () => {
	it('formats booleans through the text section only', () => {
		expect(text(true, '0.00')).toBe('TRUE');
		expect(text(false, '0.00')).toBe('FALSE');
		expect(text(true, '@')).toBe('TRUE');
		expect(text(true, '0;0;0;"<"@">"')).toBe('<TRUE>');
	});

	it('never formats errors', () => {
		expect(formatValue(ERRORS.DIV0, '0.00')).toEqual({ text: '#DIV/0!' });
		expect(formatValue(ERRORS.NA, ';;;')).toEqual({ text: '#N/A' });
	});

	it('shows blanks as empty text', () => {
		expect(formatValue(null, '0.00')).toEqual({ text: '' });
		expect(formatValue(null, '"x"@')).toEqual({ text: '' });
	});

	it('pads and fills text in the text section', () => {
		expect(formatValue('abc', '_(@_)')).toEqual({ text: ' abc ' });
		expect(formatValue('abc', '@*-')).toEqual({
			text: 'abc',
			fill: { before: 'abc', char: '-', after: '' },
		});
		expect(formatValue('abc', '*.@').fill).toEqual({ before: '', char: '.', after: 'abc' });
	});

	it('reports non-finite numbers as #NUM!', () => {
		expect(text(Number.POSITIVE_INFINITY, '0.00')).toBe('#NUM!');
	});
});

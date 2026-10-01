import { describe, expect, it } from 'vitest';

import { formatGeneral, formatValue, GENERAL_FORMAT_WIDTH } from './index';

describe('formatGeneral at the classic 11-character width', () => {
	it.each([
		[0, '0'],
		[1, '1'],
		[-1, '-1'],
		[1234.5678, '1234.5678'],
		[0.1 + 0.2, '0.3'],
		[1 / 3, '0.333333333'],
		[2 / 3, '0.666666667'],
		[Math.PI, '3.141592654'],
		[-Math.PI, '-3.141592654'],
		[100000, '100000'],
		[12345678901, '12345678901'],
		[-12345678901, '-12345678901'],
		[123456789012, '1.23457E+11'],
		[99999999999, '99999999999'],
		[99999999999.5, '1E+11'],
		[999999999999, '1E+12'],
		[1e15, '1E+15'],
		[1e21, '1E+21'],
		[1.5e300, '1.5E+300'],
		[1.7976931348623157e308, '1.7977E+308'],
		[1234567.891, '1234567.891'],
		[12345678.9012, '12345678.9'],
		[123456789.123, '123456789.1'],
		[1234567890.12, '1234567890'],
		[0.5, '0.5'],
		[0.0001, '0.0001'],
		[0.000123456789, '0.000123457'],
		[0.00001, '0.00001'],
		[0.0000015, '0.0000015'],
		[1.23456789e-5, '1.23457E-05'],
		[1e-10, '1E-10'],
		[1.5e-10, '1.5E-10'],
		[-2.5e-12, '-2.5E-12'],
		[5e-324, '4.9407E-324'],
		[0.99999999999, '1'],
		[9.99999999999, '10'],
	])('%s -> %s', (value, text) => {
		expect(formatGeneral(value)).toBe(text);
	});

	it('defaults to 11 characters', () => {
		expect(GENERAL_FORMAT_WIDTH).toBe(11);
		expect(formatGeneral(Math.PI, 0)).toBe('3.141592654');
	});

	it('reports non-finite values as #NUM!', () => {
		expect(formatGeneral(Number.NaN)).toBe('#NUM!');
		expect(formatGeneral(Number.POSITIVE_INFINITY)).toBe('#NUM!');
	});
});

describe('formatGeneral in narrower cells', () => {
	it.each([
		[Math.PI, 8, '3.141593'],
		[1 / 3, 8, '0.333333'],
		[123456789, 8, '1.23E+08'],
		[12345678, 8, '12345678'],
		[1234567.89, 8, '1234568'],
		[-1234.5678, 8, '-1234.568'],
		[123456, 5, '1E+05'],
		[0.5, 5, '0.5'],
		[1234.5678, 6, '1234.6'],
		[1e-5, 6, '1E-05'],
		[0.000123, 6, '0.0001'],
		[99.99, 3, '100'],
	])('%s in %i characters -> %s', (value, width, text) => {
		expect(formatGeneral(value, width)).toBe(text);
	});
});

describe('General through formatValue', () => {
	it('formats numbers, booleans, errors and text', () => {
		expect(formatValue(123456789012, 'General')).toEqual({ text: '1.23457E+11' });
		expect(formatValue(1234.5, 'general').text).toBe('1234.5');
		expect(formatValue(1234.5, '').text).toBe('1234.5');
		expect(formatValue(true, 'General').text).toBe('TRUE');
		expect(formatValue(false, 'General').text).toBe('FALSE');
		expect(formatValue({ type: 'error', code: '#N/A' }, 'General').text).toBe('#N/A');
		expect(formatValue('text', 'General').text).toBe('text');
		expect(formatValue(null, 'General').text).toBe('');
	});

	it('honours generalMaxChars, also inside sections', () => {
		expect(formatValue(Math.PI, 'General', { generalMaxChars: 8 }).text).toBe('3.141593');
		expect(formatValue(Math.PI, '[Red]General', { generalMaxChars: 5 })).toEqual({
			text: '3.142',
			color: '#FF0000',
		});
		expect(formatValue(-Math.PI, '"x"General', { generalMaxChars: 6 }).text).toBe('-x3.1416');
	});
});

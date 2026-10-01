import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 1234.5, A2: 0.25 });
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);

describe('TEXT', () => {
	it('formats numbers and dates with format codes', () => {
		expect(f('TEXT', 1234.5, '#,##0.00')).toBe('1,234.50');
		expect(f('TEXT', 45292, 'yyyy-mm-dd')).toBe('2024-01-01');
		expect(f('TEXT', 0.25, '0%')).toBe('25%');
		expect(f('TEXT', 1234.567, '$#,##0.00')).toBe('$1,234.57');
		expect(f('TEXT', 0.75, 'h:mm AM/PM')).toBe('6:00 PM');
		expect(f('TEXT', -5, '0;(0)')).toBe('(5)');
		expect(f('TEXT', 1234, 'General')).toBe('1234');
	});

	it('formats numeric text as a number and other text as text', () => {
		expect(f('TEXT', '1234.5', '0.0')).toBe('1234.5');
		expect(f('TEXT', '2024-01-01', 'mmm d')).toBe('Jan 1');
		expect(f('TEXT', 'abc', '0.00')).toBe('abc');
		expect(f('TEXT', 'abc', '@ "x"')).toBe('abc x');
	});

	it('handles booleans, blanks, empty formats and numeric format codes', () => {
		expect(f('TEXT', true, '0')).toBe('TRUE');
		expect(f('TEXT', null, '0.00')).toBe('0.00');
		expect(f('TEXT', 1234, '')).toBe('');
		expect(f('TEXT', 1234.5, 0)).toBe('1235');
	});

	it('propagates errors and rejects unformattable values', () => {
		expect(f('TEXT', ERRORS.NA, '0')).toBe(ERRORS.NA);
		expect(f('TEXT', 1, ERRORS.DIV0)).toBe(ERRORS.DIV0);
		expect(f('TEXT', -1, 'yyyy')).toBe(ERRORS.VALUE);
	});

	it('lifts over ranges', () => {
		expect(f('TEXT', ctx.ref('A1:A2'), '0.0')).toEqual(arr(['1234.5'], ['0.3']));
	});
});

describe('FIXED / DOLLAR', () => {
	it.each([
		[[1234.567, 1], '1,234.6'],
		[[1234.567, -1], '1,230'],
		[[-1234.567, -1, true], '-1230'],
		[[44.332], '44.33'],
		[[0.5, 0], '1'],
		[[-0.001, 2], '0.00'],
		[[1234567.891, 2, false], '1,234,567.89'],
		[[2.675, 2], '2.68'],
		[[1 / 3, 20], '0.33333333333333300000'],
		[[1e20, 2], '100,000,000,000,000,000,000.00'],
		[[0.000123, 6], '0.000123'],
		[[999.995, 2], '1,000.00'],
		[['1234.5', 0, 1], '1235'],
	] as Array<[FunctionArg[], string]>)('FIXED(%j) = %s', (args, expected) => {
		expect(f('FIXED', ...args)).toBe(expected);
	});

	it.each([
		[[1234.567, 2], '$1,234.57'],
		[[1234.567, -2], '$1,200'],
		[[-1234.567, -2], '($1,200)'],
		[[-0.123, 4], '($0.1230)'],
		[[99.888], '$99.89'],
		[[0], '$0.00'],
	] as Array<[FunctionArg[], string]>)('DOLLAR(%j) = %s', (args, expected) => {
		expect(f('DOLLAR', ...args)).toBe(expected);
	});

	it('validates arguments', () => {
		expect(f('FIXED', 1, 128)).toBe(ERRORS.VALUE);
		expect(f('DOLLAR', 1, 128)).toBe(ERRORS.VALUE);
		expect(f('FIXED', 'abc')).toBe(ERRORS.VALUE);
		expect(f('FIXED', 1, 2, 'maybe')).toBe(ERRORS.VALUE);
		expect(f('FIXED', 1, -500)).toBe('0');
		expect(f('DOLLAR', ERRORS.NA)).toBe(ERRORS.NA);
	});
});

describe('VALUE', () => {
	it('converts numeric text', () => {
		expect(f('VALUE', '$1,000')).toBe(1000);
		expect(f('VALUE', '1/1/2024')).toBe(45292);
		expect(f('VALUE', '50%')).toBe(0.5);
		expect(f('VALUE', ' 12 ')).toBe(12);
		expect(f('VALUE', '1e3')).toBe(1000);
		expect(f('VALUE', '(42)')).toBe(-42);
		expect(f('VALUE', 12)).toBe(12);
		expect(f('VALUE', null)).toBe(0);
		expect(f('VALUE', '16:48:00') as number).toBeCloseTo(0.7, 12);
	});

	it('rejects non-numeric text and booleans', () => {
		expect(f('VALUE', 'abc')).toBe(ERRORS.VALUE);
		expect(f('VALUE', '')).toBe(ERRORS.VALUE);
		expect(f('VALUE', true)).toBe(ERRORS.VALUE);
		expect(f('VALUE', ERRORS.NUM)).toBe(ERRORS.NUM);
	});
});

describe('NUMBERVALUE', () => {
	it('parses with custom separators', () => {
		expect(f('NUMBERVALUE', '2.500,27', ',', '.')).toBe(2500.27);
		expect(f('NUMBERVALUE', '3.5%')).toBe(0.035);
		expect(f('NUMBERVALUE', ' 3 000 ')).toBe(3000);
		expect(f('NUMBERVALUE', '')).toBe(0);
		expect(f('NUMBERVALUE', '9%%')).toBe(0.0009);
		expect(f('NUMBERVALUE', '1,23.4')).toBe(123.4);
		expect(f('NUMBERVALUE', '-12.5')).toBe(-12.5);
		expect(f('NUMBERVALUE', '1E3')).toBe(1000);
		expect(f('NUMBERVALUE', '.5')).toBe(0.5);
		expect(f('NUMBERVALUE', '1 234,5', ',', ' ')).toBe(1234.5);
		expect(f('NUMBERVALUE', 42)).toBe(42);
		expect(f('NUMBERVALUE', '1,5', ',,', '..')).toBe(1.5);
	});

	it('rejects malformed text and separators', () => {
		expect(f('NUMBERVALUE', '1.2.3')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', '1,000.5,0')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', 'abc')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', '$100')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', '1.5', '.', '.')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', '1.5', '')).toBe(ERRORS.VALUE);
		expect(f('NUMBERVALUE', true)).toBe(ERRORS.VALUE);
		expect(fn(TEXT_FUNCTIONS, 'NUMBERVALUE').filePrefix).toBe('_xlfn.');
	});
});

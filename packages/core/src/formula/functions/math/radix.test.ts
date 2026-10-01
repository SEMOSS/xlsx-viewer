import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';
import { fromRoman, toRoman } from './radix';

const ctx = new MockContext({ A1: true, A2: false });
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

describe('ROMAN', () => {
	it.each([
		[499, 0, 'CDXCIX'],
		[499, 1, 'LDVLIV'],
		[499, 2, 'XDIX'],
		[499, 3, 'VDIV'],
		[499, 4, 'ID'],
		[1999, 0, 'MCMXCIX'],
		[1999, 4, 'MIM'],
		[2013, 0, 'MMXIII'],
		[3999, 0, 'MMMCMXCIX'],
		[8, 0, 'VIII'],
		[45, 1, 'VL'],
		[0, 0, ''],
	])('ROMAN(%d, %d) = %s', (n, form, expected) => {
		expect(f('ROMAN', n, form)).toBe(expected);
	});

	it('maps TRUE / FALSE / omitted forms', () => {
		expect(f('ROMAN', 499)).toBe('CDXCIX');
		expect(f('ROMAN', 499, true)).toBe('CDXCIX');
		expect(f('ROMAN', 499, false)).toBe('ID');
		expect(f('ROMAN', 499, ctx.ref('A2'))).toBe('ID');
		expect(f('ROMAN', 499, null)).toBe('CDXCIX');
		expect(f('ROMAN', 1.9)).toBe('I');
	});

	it('rejects out-of-range numbers and forms', () => {
		expect(f('ROMAN', 4000)).toBe(ERRORS.VALUE);
		expect(f('ROMAN', -1)).toBe(ERRORS.VALUE);
		expect(f('ROMAN', 1, 5)).toBe(ERRORS.VALUE);
		expect(f('ROMAN', 1, -1)).toBe(ERRORS.VALUE);
		expect(f('ROMAN', 'x')).toBe(ERRORS.VALUE);
	});

	it('round-trips through ARABIC for every classic numeral', () => {
		for (let n = 1; n < 4000; n += 7) {
			expect(fromRoman(toRoman(n, 0))).toBe(n);
			expect(fromRoman(toRoman(n, 4))).toBe(n);
		}
	});
});

describe('ARABIC', () => {
	it('converts Roman numerals', () => {
		expect(f('ARABIC', 'LVII')).toBe(57);
		expect(f('ARABIC', 'mcmxii')).toBe(1912);
		expect(f('ARABIC', '')).toBe(0);
		expect(f('ARABIC', '-MMXI')).toBe(-2011);
		expect(f('ARABIC', ' XIV ')).toBe(14);
		expect(f('ARABIC', arr(['I', 'IV']))).toEqual(arr([1, 4]));
	});

	it('rejects anything else', () => {
		expect(f('ARABIC', 'ABC')).toBe(ERRORS.VALUE);
		expect(f('ARABIC', 5)).toBe(ERRORS.VALUE);
		expect(f('ARABIC', 'I'.repeat(256))).toBe(ERRORS.VALUE);
		expect(f('ARABIC', ERRORS.NA)).toBe(ERRORS.NA);
	});
});

describe('BASE / DECIMAL', () => {
	it('converts to a radix', () => {
		expect(f('BASE', 7, 2)).toBe('111');
		expect(f('BASE', 100, 16)).toBe('64');
		expect(f('BASE', 15, 2, 10)).toBe('0000001111');
		expect(f('BASE', 255, 16)).toBe('FF');
		expect(f('BASE', 35, 36)).toBe('Z');
		expect(f('BASE', 10.9, 2)).toBe('1010');
		expect(f('BASE', 0, 2)).toBe('0');
		expect(f('BASE', 2 ** 53 - 1, 36)).toBe('2GOSA7PA2GV');
	});

	it('validates BASE arguments', () => {
		expect(f('BASE', -1, 2)).toBe(ERRORS.NUM);
		expect(f('BASE', 2 ** 53, 2)).toBe(ERRORS.NUM);
		expect(f('BASE', 10, 1)).toBe(ERRORS.NUM);
		expect(f('BASE', 10, 37)).toBe(ERRORS.NUM);
		expect(f('BASE', 10, 2, 256)).toBe(ERRORS.NUM);
		expect(f('BASE', 'x', 2)).toBe(ERRORS.VALUE);
	});

	it('converts from a radix', () => {
		expect(f('DECIMAL', 'FF', 16)).toBe(255);
		expect(f('DECIMAL', 'ff', 16)).toBe(255);
		expect(f('DECIMAL', 111, 2)).toBe(7);
		expect(f('DECIMAL', 'zap', 36)).toBe(45745);
		expect(f('DECIMAL', '', 10)).toBe(0);
	});

	it('validates DECIMAL arguments', () => {
		expect(f('DECIMAL', '12', 2)).toBe(ERRORS.NUM);
		expect(f('DECIMAL', 'FF', 37)).toBe(ERRORS.NUM);
		expect(f('DECIMAL', '1 0', 10)).toBe(ERRORS.NUM);
		expect(f('DECIMAL', 'Z'.repeat(20), 36)).toBe(ERRORS.NUM);
	});
});

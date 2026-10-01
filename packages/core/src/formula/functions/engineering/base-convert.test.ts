import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);

describe('documented examples', () => {
	it.each([
		['BIN2DEC', [1100100], 100],
		['BIN2DEC', [1111111111], -1],
		['BIN2HEX', [11111011, 4], '00FB'],
		['BIN2HEX', [1110], 'E'],
		['BIN2HEX', [1111111111], 'FFFFFFFFFF'],
		['BIN2OCT', [1001, 3], '011'],
		['BIN2OCT', [1100100], '144'],
		['BIN2OCT', [1111111111], '7777777777'],
		['DEC2BIN', [9, 4], '1001'],
		['DEC2BIN', [-100], '1110011100'],
		['DEC2HEX', [100, 4], '0064'],
		['DEC2HEX', [-54], 'FFFFFFFFCA'],
		['DEC2HEX', [28], '1C'],
		['DEC2OCT', [58, 3], '072'],
		['DEC2OCT', [-100], '7777777634'],
		['HEX2BIN', ['F', 8], '00001111'],
		['HEX2BIN', ['B7'], '10110111'],
		['HEX2BIN', ['FFFFFFFFFF'], '1111111111'],
		['HEX2DEC', ['A5'], 165],
		['HEX2DEC', ['FFFFFFFF5B'], -165],
		['HEX2DEC', ['3DA408B9'], 1034160313],
		['HEX2OCT', ['F', 3], '017'],
		['HEX2OCT', ['3B4E'], '35516'],
		['HEX2OCT', ['FFFFFFFF00'], '7777777400'],
		['OCT2BIN', [3, 3], '011'],
		['OCT2BIN', [7777777000], '1000000000'],
		['OCT2DEC', [54], 44],
		['OCT2DEC', [7777777533], -165],
		['OCT2HEX', [100, 4], '0040'],
		['OCT2HEX', [7777777533], 'FFFFFFFF5B'],
	])('%s(%j) = %j', (name, args, expected) => {
		expect(f(name, ...(args as FunctionArg[]))).toBe(expected);
	});
});

describe("ranges and two's complement", () => {
	it('accepts the full signed range of each base', () => {
		expect(f('DEC2BIN', 511)).toBe('111111111');
		expect(f('DEC2BIN', -512)).toBe('1000000000');
		expect(f('DEC2OCT', 536870911)).toBe('3777777777');
		expect(f('DEC2OCT', -536870912)).toBe('4000000000');
		expect(f('DEC2HEX', 549755813887)).toBe('7FFFFFFFFF');
		expect(f('DEC2HEX', -549755813888)).toBe('8000000000');
		expect(f('BIN2DEC', '1000000000')).toBe(-512);
		expect(f('OCT2DEC', '4000000000')).toBe(-536870912);
		expect(f('HEX2DEC', '8000000000')).toBe(-549755813888);
		expect(f('HEX2DEC', '7FFFFFFFFF')).toBe(549755813887);
	});

	it('rejects values outside the target range', () => {
		expect(f('DEC2BIN', 512)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', -513)).toBe(ERRORS.NUM);
		expect(f('DEC2OCT', 536870912)).toBe(ERRORS.NUM);
		expect(f('DEC2HEX', 549755813888)).toBe(ERRORS.NUM);
		expect(f('HEX2BIN', '200')).toBe(ERRORS.NUM);
		expect(f('HEX2BIN', 'FFFFFFFDFF')).toBe(ERRORS.NUM);
		expect(f('HEX2BIN', 'FFFFFFFE00')).toBe('1000000000');
		expect(f('OCT2BIN', '1000')).toBe(ERRORS.NUM);
		expect(f('HEX2OCT', '20000000')).toBe(ERRORS.NUM);
		expect(f('HEX2OCT', 'FFE0000000')).toBe('4000000000');
	});

	it('round-trips through every base', () => {
		for (const n of [0, 1, 7, 100, 255, 511, -1, -100, -512]) {
			expect(f('BIN2DEC', f('DEC2BIN', n) as string)).toBe(n);
			expect(f('OCT2DEC', f('DEC2OCT', n) as string)).toBe(n);
			expect(f('HEX2DEC', f('DEC2HEX', n) as string)).toBe(n);
			expect(f('HEX2BIN', f('BIN2HEX', f('DEC2BIN', n) as string) as string)).toBe(f('DEC2BIN', n));
			expect(f('OCT2HEX', f('HEX2OCT', f('DEC2HEX', n) as string) as string)).toBe(f('DEC2HEX', n));
		}
	});
});

describe('places', () => {
	it('pads with zeros and is truncated', () => {
		expect(f('DEC2BIN', 9, 8)).toBe('00001001');
		expect(f('DEC2BIN', 9, 4.9)).toBe('1001');
		expect(f('DEC2HEX', 255, 10)).toBe('00000000FF');
		expect(f('OCT2BIN', 7, 5)).toBe('00111');
	});

	it('is ignored for negative numbers but still validated', () => {
		expect(f('DEC2BIN', -1, 2)).toBe('1111111111');
		expect(f('DEC2HEX', -1, 3)).toBe('FFFFFFFFFF');
		expect(f('DEC2BIN', -1, 11)).toBe(ERRORS.NUM);
	});

	it('rejects values that are too small, too large or not numeric', () => {
		expect(f('DEC2HEX', 64, 1)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', 9, 3)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', 9, 0)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', 9, -1)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', 9, 11)).toBe(ERRORS.NUM);
		expect(f('DEC2BIN', 9, 'x')).toBe(ERRORS.VALUE);
		expect(f('DEC2BIN', 9, null)).toBe('1001');
	});
});

describe('input validation', () => {
	it('accepts text and numbers, case-insensitive hex digits', () => {
		expect(f('HEX2DEC', 'ff')).toBe(255);
		expect(f('HEX2DEC', 100)).toBe(256);
		expect(f('BIN2DEC', '0000000001')).toBe(1);
		expect(f('DEC2BIN', '9')).toBe('1001');
		expect(f('DEC2BIN', 9.99)).toBe('1001');
		expect(f('DEC2BIN', -0.5)).toBe('0');
		expect(f('BIN2DEC', '')).toBe(0);
		expect(f('BIN2DEC', null)).toBe(0);
	});

	it('rejects invalid digits, long numbers and logical values', () => {
		expect(f('BIN2DEC', 1102)).toBe(ERRORS.NUM);
		expect(f('BIN2DEC', '10000000000')).toBe(ERRORS.NUM);
		expect(f('BIN2DEC', -1)).toBe(ERRORS.NUM);
		expect(f('BIN2DEC', 1.5)).toBe(ERRORS.NUM);
		expect(f('OCT2DEC', '8')).toBe(ERRORS.NUM);
		expect(f('HEX2DEC', 'G1')).toBe(ERRORS.NUM);
		expect(f('HEX2DEC', ' A')).toBe(ERRORS.NUM);
		expect(f('HEX2DEC', 1e10)).toBe(ERRORS.NUM);
		expect(f('BIN2DEC', true)).toBe(ERRORS.VALUE);
		expect(f('DEC2BIN', true)).toBe(ERRORS.VALUE);
		expect(f('DEC2BIN', 'nine')).toBe(ERRORS.VALUE);
		expect(f('HEX2DEC', ERRORS.REF)).toBe(ERRORS.REF);
	});

	it('lifts over arrays', () => {
		expect(f('DEC2HEX', arr([10, 11, 12]))).toEqual(arr(['A', 'B', 'C']));
	});
});

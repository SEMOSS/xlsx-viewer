import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../values/errors';
import type { ArrayValue, Scalar } from '../types';
import {
	applyOperator,
	broadcastBinary,
	compareNumbers,
	mapArray,
	negate,
	numberResult,
	percent,
} from './operators';

const arr = (...rows: Scalar[][]): ArrayValue => ({
	kind: 'array',
	rows: rows.length,
	cols: rows[0].length,
	values: rows,
});

describe('arithmetic', () => {
	it('coerces blanks, booleans and numeric text', () => {
		expect(applyOperator('+', null, 2)).toBe(2);
		expect(applyOperator('+', true, true)).toBe(2);
		expect(applyOperator('*', '3', '4')).toBe(12);
		expect(applyOperator('-', ' 10 ', 1)).toBe(9);
		expect(applyOperator('+', '1/2/2024', 0)).toBeGreaterThan(45000);
		expect(applyOperator('*', '50%', 4)).toBe(2);
	});

	it('rejects non-numeric text with #VALUE!', () => {
		expect(applyOperator('+', 'abc', 1)).toBe(ERRORS.VALUE);
		expect(applyOperator('+', '', 1)).toBe(ERRORS.VALUE);
	});

	it('propagates errors left to right', () => {
		expect(applyOperator('+', ERRORS.NA, ERRORS.DIV0)).toBe(ERRORS.NA);
		expect(applyOperator('+', 1, ERRORS.REF)).toBe(ERRORS.REF);
		expect(applyOperator('*', 'x', ERRORS.NA)).toBe(ERRORS.VALUE);
	});

	it('handles division and powers like Excel', () => {
		expect(applyOperator('/', 1, 0)).toBe(ERRORS.DIV0);
		expect(applyOperator('/', 1, null)).toBe(ERRORS.DIV0);
		expect(applyOperator('^', 0, 0)).toBe(ERRORS.NUM);
		expect(applyOperator('^', 0, -1)).toBe(ERRORS.DIV0);
		expect(applyOperator('^', -8, 1 / 3)).toBe(ERRORS.NUM);
		expect(applyOperator('^', -2, 3)).toBe(-8);
		expect(applyOperator('^', 2, 0.5)).toBeCloseTo(Math.SQRT2, 15);
		expect(applyOperator('^', 10, 400)).toBe(ERRORS.NUM);
		expect(applyOperator('*', 1e308, 10)).toBe(ERRORS.NUM);
	});

	it('normalises negative zero', () => {
		expect(Object.is(applyOperator('*', -1, 0), 0)).toBe(true);
		expect(Object.is(negate(0), 0)).toBe(true);
		expect(numberResult(Number.NaN)).toBe(ERRORS.NUM);
	});
});

describe('unary operators', () => {
	it('negates and takes percentages with coercion', () => {
		expect(negate('5')).toBe(-5);
		expect(negate(true)).toBe(-1);
		expect(negate(null)).toBe(0);
		expect(negate('a')).toBe(ERRORS.VALUE);
		expect(percent(50)).toBe(0.5);
		expect(percent('10')).toBe(0.1);
		expect(percent(ERRORS.NA)).toBe(ERRORS.NA);
	});
});

describe('text concatenation', () => {
	it('joins with Excel number formatting', () => {
		expect(applyOperator('&', 'a', 1.5)).toBe('a1.5');
		expect(applyOperator('&', 0.1 + 0.2, '')).toBe('0.3');
		expect(applyOperator('&', true, null)).toBe('TRUE');
		expect(applyOperator('&', 1e21, '')).toBe('1E+21');
		expect(applyOperator('&', ERRORS.NA, 'x')).toBe(ERRORS.NA);
	});

	it('limits text to 32767 characters', () => {
		expect(applyOperator('&', 'a'.repeat(32767), 'b')).toBe(ERRORS.VALUE);
		expect((applyOperator('&', 'a'.repeat(32766), 'b') as string).length).toBe(32767);
	});
});

describe('comparisons', () => {
	it('orders numbers < text < booleans and ignores case', () => {
		expect(applyOperator('<', 1e9, 'a')).toBe(true);
		expect(applyOperator('<', 'zzz', false)).toBe(true);
		expect(applyOperator('=', 'abc', 'ABC')).toBe(true);
		expect(applyOperator('<>', 'abc', 'abd')).toBe(true);
		expect(applyOperator('>=', 'b', 'A')).toBe(true);
	});

	it('treats blank as 0, "" or FALSE depending on the other side', () => {
		expect(applyOperator('=', null, 0)).toBe(true);
		expect(applyOperator('=', null, '')).toBe(true);
		expect(applyOperator('=', null, false)).toBe(true);
		expect(applyOperator('=', null, null)).toBe(true);
		expect(applyOperator('<', null, 1)).toBe(true);
	});

	it('does not coerce text to numbers', () => {
		expect(applyOperator('=', '1', 1)).toBe(false);
		expect(applyOperator('>', '1', 2)).toBe(true);
	});

	it('compares numbers to 15 significant digits', () => {
		expect(applyOperator('=', 0.1 + 0.2, 0.3)).toBe(true);
		expect(applyOperator('<', 0.1 + 0.2, 0.3)).toBe(false);
		expect(applyOperator('=', 1, 1.000000000001)).toBe(false);
		expect(compareNumbers(1e-20, 2e-20)).toBe(-1);
		expect(compareNumbers(5, 5)).toBe(0);
	});

	it('propagates errors', () => {
		expect(applyOperator('=', ERRORS.NA, 1)).toBe(ERRORS.NA);
		expect(applyOperator('=', 1, ERRORS.DIV0)).toBe(ERRORS.DIV0);
	});
});

describe('broadcasting', () => {
	const add = (x: Scalar, y: Scalar) => applyOperator('+', x, y);

	it('combines arrays of the same shape element-wise', () => {
		expect(broadcastBinary(arr([1, 2], [3, 4]), arr([10, 20], [30, 40]), add).values).toEqual([
			[11, 22],
			[33, 44],
		]);
	});

	it('repeats scalars, single rows and single columns', () => {
		expect(broadcastBinary(arr([1, 2, 3]), 10, add).values).toEqual([[11, 12, 13]]);
		expect(broadcastBinary(arr([1, 2, 3]), arr([10], [20]), add).values).toEqual([
			[11, 12, 13],
			[21, 22, 23],
		]);
	});

	it('fills positions missing from the smaller array with #N/A', () => {
		expect(broadcastBinary(arr([1, 2]), arr([1, 2, 3]), add).values).toEqual([[2, 4, ERRORS.NA]]);
		expect(
			broadcastBinary(arr([1, 2], [3, 4]), arr([1, 1, 1], [1, 1, 1], [1, 1, 1]), add).values,
		).toEqual([
			[2, 3, ERRORS.NA],
			[4, 5, ERRORS.NA],
			[ERRORS.NA, ERRORS.NA, ERRORS.NA],
		]);
	});

	it('maps unary functions over arrays', () => {
		expect(mapArray(arr([1, -2], ['3', 'x']), (v) => negate(v)).values).toEqual([
			[-1, 2],
			[-3, ERRORS.VALUE],
		]);
	});
});

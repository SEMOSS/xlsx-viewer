import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';
import {
	approxFloor,
	cleanNoise,
	roundDirected,
	roundHalfAway,
	shiftDecimal,
	snapToInteger,
} from './rounding';

const ctx = new MockContext({ A1: 1.25, A2: -1.35, A3: 'x', B1: 2 });
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

describe('precision helpers', () => {
	it('shifts decimals exactly', () => {
		expect(shiftDecimal(2.675, 2)).toBe(267.5);
		expect(shiftDecimal(1.005, 2)).toBe(100.5);
		expect(shiftDecimal(1234, -2)).toBe(12.34);
		expect(shiftDecimal(1e-7, 7)).toBe(1);
	});

	it('snaps values within binary noise of an integer', () => {
		expect(snapToInteger(2.9999999999999996)).toBe(3);
		expect(snapToInteger(3.0000000000000004)).toBe(3);
		expect(snapToInteger(2.9999)).toBe(2.9999);
		expect(snapToInteger(1e-17)).toBe(1e-17);
		expect(approxFloor((0.1 + 0.7) * 10)).toBe(8);
	});

	it('cleans noise but keeps genuine digits', () => {
		expect(cleanNoise(0.1 + 0.2)).toBe(0.3);
		expect(cleanNoise(7 * 0.2)).toBe(1.4);
		expect(cleanNoise(Math.LN2)).toBe(Math.LN2);
		expect(cleanNoise(1234567890123.4567)).toBe(1234567890123.4567);
	});

	it('rounds half away from zero on the 15-digit form', () => {
		expect(roundHalfAway(2.675, 2)).toBe(2.68);
		expect(roundHalfAway(1.005, 2)).toBe(1.01);
		expect(roundHalfAway(0.285, 2)).toBe(0.29);
		expect(roundHalfAway(-0.5, 0)).toBe(-1);
		expect(roundHalfAway(1e300, 2)).toBe(1e300);
		expect(roundHalfAway(123.456, 400)).toBe(123.456);
		expect(roundHalfAway(5e-324, 2)).toBe(0);
		expect(Object.is(roundHalfAway(-0.4, 0), 0)).toBe(true);
	});

	it('rounds toward / away from zero', () => {
		expect(roundDirected(0.1 + 0.2, 1, true)).toBe(0.3);
		expect(roundDirected(2.2 * 100, 0, true)).toBe(220);
		expect(roundDirected(123456789012345.67, 0, false)).toBe(123456789012345);
		expect(roundDirected(5e-324, 0, true)).toBe(1);
	});
});

describe('ROUND / ROUNDUP / ROUNDDOWN', () => {
	it.each([
		[2.15, 1, 2.2],
		[2.149, 1, 2.1],
		[-1.475, 2, -1.48],
		[21.5, -1, 20],
		[626.3, -3, 1000],
		[1.98, -1, 0],
		[-50.55, -2, -100],
		[2.675, 2, 2.68],
		[0.5, 0, 1],
		[-0.5, 0, -1],
		[2.5, 0, 3],
		[-2.5, 0, -3],
		[1234.5678, 2.9, 1234.57],
		[1234.5678, -1.5, 1230],
	])('ROUND(%d, %d) = %d', (x, d, expected) => {
		expect(f('ROUND', x, d)).toBe(expected);
	});

	it.each([
		[3.2, 0, 4],
		[76.9, 0, 77],
		[3.14159, 3, 3.142],
		[-3.14159, 1, -3.2],
		[31415.92654, -2, 31500],
		[4, -2, 100],
	])('ROUNDUP(%d, %d) = %d', (x, d, expected) => {
		expect(f('ROUNDUP', x, d)).toBe(expected);
	});

	it.each([
		[3.2, 0, 3],
		[76.9, 0, 76],
		[3.14159, 3, 3.141],
		[-3.14159, 1, -3.1],
		[31415.92654, -2, 31400],
		[0.1 + 0.2, 1, 0.3],
	])('ROUNDDOWN(%d, %d) = %d', (x, d, expected) => {
		expect(f('ROUNDDOWN', x, d)).toBe(expected);
	});

	it('coerces text, booleans and blanks', () => {
		expect(f('ROUND', '2.5', 0)).toBe(3);
		expect(f('ROUND', true, 0)).toBe(1);
		expect(f('ROUND', null, 0)).toBe(0);
		expect(f('ROUND', 1.5, null)).toBe(2);
		expect(f('ROUND', '1.25', '1')).toBe(1.3);
	});

	it('propagates errors and rejects bad input', () => {
		expect(f('ROUND', 'abc', 0)).toBe(ERRORS.VALUE);
		expect(f('ROUND', ERRORS.DIV0, 0)).toBe(ERRORS.DIV0);
		expect(f('ROUND', 1, ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('ROUND', 1.5)).toBe(ERRORS.VALUE);
		expect(f('ROUNDUP', ctx.ref('A3'), 0)).toBe(ERRORS.VALUE);
	});

	it('lifts over ranges and arrays', () => {
		expect(f('ROUND', ctx.ref('A1:A2'), 1)).toEqual(arr([1.3], [-1.4]));
		expect(f('ROUNDUP', arr([1.21, 1.29]), 1)).toEqual(arr([1.3, 1.3]));
		expect(f('ROUNDDOWN', 1.987, arr([0, 1, 2]))).toEqual(arr([1, 1.9, 1.98]));
	});
});

describe('TRUNC / INT', () => {
	it('truncates toward zero', () => {
		expect(f('TRUNC', 8.9)).toBe(8);
		expect(f('TRUNC', -8.9)).toBe(-8);
		expect(f('TRUNC', 0.45)).toBe(0);
		expect(f('TRUNC', 3.14159, 2)).toBe(3.14);
		expect(f('TRUNC', 1234.5, -2)).toBe(1200);
		expect(f('TRUNC', 'x')).toBe(ERRORS.VALUE);
	});

	it('floors toward negative infinity', () => {
		expect(f('INT', 8.9)).toBe(8);
		expect(f('INT', -8.9)).toBe(-9);
		expect(f('INT', -0.5)).toBe(-1);
		expect(f('INT', (0.1 + 0.7) * 10)).toBe(8);
		expect(f('INT', '5.5')).toBe(5);
		expect(f('INT', true)).toBe(1);
	});
});

describe('EVEN / ODD', () => {
	it.each([
		[1.5, 2],
		[3, 4],
		[2, 2],
		[-1, -2],
		[0, 0],
		[-2.5, -4],
	])('EVEN(%d) = %d', (x, expected) => {
		expect(f('EVEN', x)).toBe(expected);
	});

	it.each([
		[1.5, 3],
		[3, 3],
		[2, 3],
		[-1, -1],
		[-2, -3],
		[0, 1],
	])('ODD(%d) = %d', (x, expected) => {
		expect(f('ODD', x)).toBe(expected);
	});
});

describe('MROUND', () => {
	it('rounds to the nearest multiple', () => {
		expect(f('MROUND', 10, 3)).toBe(9);
		expect(f('MROUND', -10, -3)).toBe(-9);
		expect(f('MROUND', 1.3, 0.2)).toBe(1.4);
		expect(f('MROUND', 6.05, 0.1)).toBe(6.1);
		expect(f('MROUND', 7.5, 5)).toBe(10);
		expect(f('MROUND', 5, 0)).toBe(0);
	});

	it('requires matching signs', () => {
		expect(f('MROUND', 5, -2)).toBe(ERRORS.NUM);
		expect(f('MROUND', -5, 2)).toBe(ERRORS.NUM);
	});
});

describe('CEILING / FLOOR', () => {
	it.each([
		[2.5, 1, 3],
		[-2.5, -2, -4],
		[-2.5, 2, -2],
		[1.5, 0.1, 1.5],
		[0.234, 0.01, 0.24],
		[0.3, 0.1, 0.3],
		[5, 0, 0],
		[0, -1, 0],
	])('CEILING(%d, %d) = %d', (x, s, expected) => {
		expect(f('CEILING', x, s)).toBe(expected);
	});

	it.each([
		[3.7, 2, 2],
		[-2.5, -2, -2],
		[1.58, 0.1, 1.5],
		[0.234, 0.01, 0.23],
		[-2.5, 2, -4],
		[0, 0, 0],
	])('FLOOR(%d, %d) = %d', (x, s, expected) => {
		expect(f('FLOOR', x, s)).toBe(expected);
	});

	it('reports invalid sign combinations and zero significance', () => {
		expect(f('CEILING', 2.5, -2)).toBe(ERRORS.NUM);
		expect(f('FLOOR', 2.5, -2)).toBe(ERRORS.NUM);
		expect(f('FLOOR', 5, 0)).toBe(ERRORS.DIV0);
		expect(f('CEILING', 2.5)).toBe(ERRORS.VALUE);
	});

	it.each([
		['CEILING.MATH', [24.3, 5], 25],
		['CEILING.MATH', [6.7], 7],
		['CEILING.MATH', [-8.1, 2], -8],
		['CEILING.MATH', [-5.5, 2, -1], -6],
		['CEILING.MATH', [-5.5, -2], -4],
		['CEILING.MATH', [5.5, 0], 0],
		['FLOOR.MATH', [24.3, 5], 20],
		['FLOOR.MATH', [6.7], 6],
		['FLOOR.MATH', [-8.1, 2], -10],
		['FLOOR.MATH', [-5.5, 2, -1], -4],
		['CEILING.PRECISE', [4.3], 5],
		['CEILING.PRECISE', [-4.3], -4],
		['CEILING.PRECISE', [4.3, 2], 6],
		['CEILING.PRECISE', [4.3, -2], 6],
		['CEILING.PRECISE', [-4.3, 2], -4],
		['CEILING.PRECISE', [-4.3, -2], -4],
		['ISO.CEILING', [4.3], 5],
		['ISO.CEILING', [-4.3, 2], -4],
		['FLOOR.PRECISE', [-3.2, -1], -4],
		['FLOOR.PRECISE', [3.2, 1], 3],
		['FLOOR.PRECISE', [-3.2, 1], -4],
		['FLOOR.PRECISE', [3.2, -1], 3],
		['FLOOR.PRECISE', [3.2], 3],
	] as Array<[string, number[], number]>)('%s(%j) = %d', (name, args, expected) => {
		expect(f(name, ...args)).toBe(expected);
	});

	it('uses _xlfn. prefixes for the 2010+ variants only', () => {
		expect(fn(MATH_FUNCTIONS, 'CEILING.MATH').filePrefix).toBe('_xlfn.');
		expect(fn(MATH_FUNCTIONS, 'FLOOR.PRECISE').filePrefix).toBe('_xlfn.');
		expect(fn(MATH_FUNCTIONS, 'ISO.CEILING').filePrefix).toBeUndefined();
		expect(fn(MATH_FUNCTIONS, 'CEILING').filePrefix).toBeUndefined();
	});
});

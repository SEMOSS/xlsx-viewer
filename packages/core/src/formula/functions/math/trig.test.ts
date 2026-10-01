import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

describe('circular functions', () => {
	it.each([
		['SIN', Math.PI / 2, 1],
		['SIN', (30 * Math.PI) / 180, 0.5],
		['COS', 1.047, 0.5001711],
		['COS', 0, 1],
		['TAN', 0.785, 0.99920399],
		['TAN', (45 * Math.PI) / 180, 1],
		['ASIN', -0.5, -0.523598776],
		['ACOS', -0.5, 2.094395102],
		['ATAN', 1, 0.785398163],
		['COT', 30, -0.156119952],
		['CSC', 15, 1.537780562],
		['SEC', 45, 1.903594407],
		['SEC', 0, 1],
		['ACOT', 2, 0.463647609],
		['ACOT', 0, Math.PI / 2],
		['ACOT', -1, (3 * Math.PI) / 4],
	] as Array<[string, number, number]>)('%s(%d) = %d', (name, x, expected) => {
		expect(f(name, x)).toBeCloseTo(expected, 7);
	});

	it('returns errors outside the domain', () => {
		expect(f('ASIN', 2)).toBe(ERRORS.NUM);
		expect(f('ACOS', -1.5)).toBe(ERRORS.NUM);
		expect(f('COT', 0)).toBe(ERRORS.DIV0);
		expect(f('CSC', 0)).toBe(ERRORS.DIV0);
		expect(f('COT', 2 ** 27)).toBe(ERRORS.NUM);
		expect(f('CSC', -(2 ** 27))).toBe(ERRORS.NUM);
		expect(f('SEC', 2 ** 27)).toBe(ERRORS.NUM);
		expect(f('SIN', 'x')).toBe(ERRORS.VALUE);
	});

	it('takes ATAN2 arguments as (x, y)', () => {
		expect(f('ATAN2', 1, 1)).toBeCloseTo(0.785398163, 9);
		expect(f('ATAN2', -1, -1)).toBeCloseTo(-2.35619449, 8);
		expect(f('ATAN2', 1, 0)).toBe(0);
		expect(f('ATAN2', 0, 1)).toBe(Math.PI / 2);
		expect(f('ATAN2', 0, 0)).toBe(ERRORS.DIV0);
	});

	it('lifts over arrays', () => {
		const result = f('COS', arr([0, Math.PI]));
		expect(result).toEqual(arr([1, -1]));
	});
});

describe('hyperbolic functions', () => {
	it.each([
		['SINH', 1, 1.175201194],
		['COSH', 4, 27.30823284],
		['COSH', Math.E, 7.610125139],
		['TANH', -2, -0.96402758],
		['TANH', 0.5, 0.462117157],
		['ASINH', -2.5, -1.647231146],
		['ASINH', 10, 2.99822295],
		['ACOSH', 1, 0],
		['ACOSH', 10, 2.993222846],
		['ATANH', 0.76159416, 1],
		['ATANH', -0.1, -0.100335348],
		['COTH', 2, 1.037314721],
		['CSCH', 1.5, 0.4696424406],
		['SECH', 0, 1],
		['ACOTH', 6, 0.168236118],
		['ACOTH', -6, -0.168236118],
	] as Array<[string, number, number]>)('%s(%d) = %d', (name, x, expected) => {
		expect(f(name, x)).toBeCloseTo(expected, 7);
	});

	it('returns errors outside the domain or on overflow', () => {
		expect(f('SINH', 1000)).toBe(ERRORS.NUM);
		expect(f('ACOSH', 0.5)).toBe(ERRORS.NUM);
		expect(f('ATANH', 1)).toBe(ERRORS.NUM);
		expect(f('ATANH', -1)).toBe(ERRORS.NUM);
		expect(f('COTH', 0)).toBe(ERRORS.DIV0);
		expect(f('CSCH', 0)).toBe(ERRORS.DIV0);
		expect(f('COTH', 2 ** 27)).toBe(ERRORS.NUM);
		expect(f('ACOTH', 1)).toBe(ERRORS.NUM);
		expect(f('ACOTH', 0.5)).toBe(ERRORS.NUM);
		expect(f('SECH', 45)).toBeCloseTo(5.73e-20, 21);
	});
});

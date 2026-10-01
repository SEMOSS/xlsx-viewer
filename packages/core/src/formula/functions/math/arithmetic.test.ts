import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: -2, A2: 3, A3: 'abc', A4: '4', A5: true, B1: ERRORS.NA });
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

describe('ABS / SIGN', () => {
	it('works on numbers and coerced values', () => {
		expect(f('ABS', -2)).toBe(2);
		expect(f('ABS', 2)).toBe(2);
		expect(f('ABS', '-3')).toBe(3);
		expect(f('ABS', true)).toBe(1);
		expect(f('ABS', null)).toBe(0);
		expect(f('ABS', ctx.ref('A4'))).toBe(4);
		expect(f('SIGN', 10)).toBe(1);
		expect(f('SIGN', 4 - 4)).toBe(0);
		expect(f('SIGN', -0.00001)).toBe(-1);
	});

	it('propagates errors and rejects text', () => {
		expect(f('ABS', 'abc')).toBe(ERRORS.VALUE);
		expect(f('ABS', ctx.ref('B1'))).toBe(ERRORS.NA);
		expect(f('SIGN', ERRORS.DIV0)).toBe(ERRORS.DIV0);
		expect(f('ABS')).toBe(ERRORS.VALUE);
		expect(f('ABS', 1, 2)).toBe(ERRORS.VALUE);
	});

	it('lifts element-wise over ranges and arrays', () => {
		expect(f('ABS', ctx.ref('A1:A3'))).toEqual(arr([2], [3], [ERRORS.VALUE]));
		expect(f('SIGN', arr([-5, 0, 5]))).toEqual(arr([-1, 0, 1]));
		expect(f('ABS', ctx.ref('A1:B1'))).toEqual(arr([2, ERRORS.NA]));
	});
});

describe('SQRT / SQRTPI / EXP / LN / LOG / LOG10', () => {
	it('computes roots and exponentials', () => {
		expect(f('SQRT', 16)).toBe(4);
		expect(f('SQRT', -16)).toBe(ERRORS.NUM);
		expect(f('SQRTPI', 1)).toBeCloseTo(1.772453851, 9);
		expect(f('SQRTPI', 2)).toBeCloseTo(2.506628275, 9);
		expect(f('SQRTPI', -1)).toBe(ERRORS.NUM);
		expect(f('EXP', 1)).toBe(Math.E);
		expect(f('EXP', 2)).toBeCloseTo(7.389056099, 9);
		expect(f('EXP', 710)).toBe(ERRORS.NUM);
	});

	it('computes logarithms', () => {
		expect(f('LN', 86)).toBeCloseTo(4.454347296, 9);
		expect(f('LN', Math.exp(3))).toBe(3);
		expect(f('LN', 0)).toBe(ERRORS.NUM);
		expect(f('LN', -1)).toBe(ERRORS.NUM);
		expect(f('LOG', 10)).toBe(1);
		expect(f('LOG', 8, 2)).toBe(3);
		expect(f('LOG', 9, 3)).toBe(2);
		expect(f('LOG', 1000)).toBe(3);
		expect(f('LOG', 86, 2.7182818)).toBeCloseTo(4.4543473, 7);
		expect(f('LOG10', 86)).toBeCloseTo(1.934498451, 9);
		expect(f('LOG10', 1e5)).toBe(5);
		expect(f('LOG10', 0)).toBe(ERRORS.NUM);
	});

	it('validates logarithm bases', () => {
		expect(f('LOG', 10, 1)).toBe(ERRORS.DIV0);
		expect(f('LOG', 0)).toBe(ERRORS.NUM);
		expect(f('LOG', 10, 0)).toBe(ERRORS.NUM);
		expect(f('LOG', 10, -2)).toBe(ERRORS.NUM);
		expect(f('LOG', 100, null)).toBe(ERRORS.NUM);
		expect(f('LOG', -1, 1)).toBe(ERRORS.NUM);
	});
});

describe('POWER', () => {
	it('raises to powers', () => {
		expect(f('POWER', 5, 2)).toBe(25);
		expect(f('POWER', 98.6, 3.2)).toBeCloseTo(2401077.222, 3);
		expect(f('POWER', 4, 5 / 4)).toBeCloseTo(5.656854249, 9);
		expect(f('POWER', -2, 3)).toBe(-8);
		expect(f('POWER', 2, -1)).toBe(0.5);
		expect(f('POWER', '2', '10')).toBe(1024);
	});

	it('handles zero and negative bases like Excel', () => {
		expect(f('POWER', 0, 0)).toBe(ERRORS.NUM);
		expect(f('POWER', 0, -1)).toBe(ERRORS.DIV0);
		expect(f('POWER', 0, 2)).toBe(0);
		expect(f('POWER', -8, 1 / 3)).toBeCloseTo(-2, 12);
		expect(f('POWER', -32, 0.2)).toBeCloseTo(-2, 12);
		expect(f('POWER', -8, 2 / 3)).toBe(ERRORS.NUM);
		expect(f('POWER', -4, 0.5)).toBe(ERRORS.NUM);
		expect(f('POWER', 10, 400)).toBe(ERRORS.NUM);
	});
});

describe('MOD / QUOTIENT', () => {
	it.each([
		[3, 2, 1],
		[-3, 2, 1],
		[3, -2, -1],
		[-3, -2, -1],
		[10.5, 3, 1.5],
		[-10.5, 3, 1.5],
		[1.1, 1, 0.1],
		[0.3, 0.1, 0],
		[5.5, 0.5, 0],
		[-0.5, 0.3, 0.1],
		[6, 3, 0],
		[12345.6, 0.1, 0],
		[-1e-20, 3, 0],
		[1e20, 3, 1],
		[2 ** 40 + 1, 2, 1],
	])('MOD(%d, %d) = %d', (n, d, expected) => {
		expect(f('MOD', n, d)).toBe(expected);
	});

	it('reports division by zero and coerces arguments', () => {
		expect(f('MOD', 3, 0)).toBe(ERRORS.DIV0);
		expect(f('MOD', 3, null)).toBe(ERRORS.DIV0);
		expect(f('MOD', '7', true)).toBe(0);
		expect(f('MOD', ctx.ref('A1:A2'), 2)).toEqual(arr([0], [1]));
	});

	it('computes integer quotients', () => {
		expect(f('QUOTIENT', 5, 2)).toBe(2);
		expect(f('QUOTIENT', 4.5, 3.1)).toBe(1);
		expect(f('QUOTIENT', -10, 3)).toBe(-3);
		expect(f('QUOTIENT', 1, 0)).toBe(ERRORS.DIV0);
		expect(f('QUOTIENT', 'a', 1)).toBe(ERRORS.VALUE);
	});
});

describe('PI / DEGREES / RADIANS', () => {
	it('converts angles', () => {
		expect(f('PI')).toBe(Math.PI);
		expect(f('PI', 1)).toBe(ERRORS.VALUE);
		expect(f('DEGREES', Math.PI)).toBe(180);
		expect(f('RADIANS', 270)).toBeCloseTo(4.71238898, 8);
		expect(f('RADIANS', 180)).toBe(Math.PI);
	});
});

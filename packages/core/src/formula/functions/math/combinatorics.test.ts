import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

const ctx = new MockContext({
	A1: 12,
	A2: 18,
	A3: 24,
	B1: 12,
	B2: 'text',
	C1: 4,
	C3: 6,
	D1: ERRORS.DIV0,
});
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

describe('FACT / FACTDOUBLE', () => {
	it('computes factorials of truncated arguments', () => {
		expect(f('FACT', 5)).toBe(120);
		expect(f('FACT', 1.9)).toBe(1);
		expect(f('FACT', 0)).toBe(1);
		expect(f('FACT', 1)).toBe(1);
		expect(f('FACT', '4')).toBe(24);
		expect(f('FACT', 170)).toBeCloseTo(7.25741561530799e306, -293);
	});

	it('rejects negatives and overflow', () => {
		expect(f('FACT', -1)).toBe(ERRORS.NUM);
		expect(f('FACT', 171)).toBe(ERRORS.NUM);
		expect(f('FACT', 'x')).toBe(ERRORS.VALUE);
	});

	it('computes double factorials', () => {
		expect(f('FACTDOUBLE', 6)).toBe(48);
		expect(f('FACTDOUBLE', 7)).toBe(105);
		expect(f('FACTDOUBLE', 7.9)).toBe(105);
		expect(f('FACTDOUBLE', 0)).toBe(1);
		expect(f('FACTDOUBLE', 1)).toBe(1);
		expect(f('FACTDOUBLE', -1)).toBe(ERRORS.NUM);
		expect(f('FACTDOUBLE', 400)).toBe(ERRORS.NUM);
	});

	it('lifts over arrays', () => {
		expect(f('FACT', arr([3, 4, -1]))).toEqual(arr([6, 24, ERRORS.NUM]));
	});
});

describe('COMBIN / COMBINA', () => {
	it('counts combinations', () => {
		expect(f('COMBIN', 8, 2)).toBe(28);
		expect(f('COMBIN', 8.9, 2.1)).toBe(28);
		expect(f('COMBIN', 5, 0)).toBe(1);
		expect(f('COMBIN', 5, 5)).toBe(1);
		expect(f('COMBIN', 49, 6)).toBe(13983816);
		expect(f('COMBIN', 100, 50)).toBeCloseTo(1.00891344545564e29, -16);
	});

	it('validates arguments', () => {
		expect(f('COMBIN', 2, 3)).toBe(ERRORS.NUM);
		expect(f('COMBIN', -1, 0)).toBe(ERRORS.NUM);
		expect(f('COMBIN', 10, -1)).toBe(ERRORS.NUM);
		expect(f('COMBIN', 2000, 1000)).toBe(ERRORS.NUM);
		expect(f('COMBIN', 'a', 1)).toBe(ERRORS.VALUE);
	});

	it('counts combinations with repetition', () => {
		expect(f('COMBINA', 4, 3)).toBe(20);
		expect(f('COMBINA', 10, 3)).toBe(220);
		expect(f('COMBINA', 0, 0)).toBe(1);
		expect(f('COMBINA', 3, 0)).toBe(1);
		expect(f('COMBINA', -1, 2)).toBe(ERRORS.NUM);
		expect(f('COMBINA', 2, -1)).toBe(ERRORS.NUM);
		expect(fn(MATH_FUNCTIONS, 'COMBINA').filePrefix).toBe('_xlfn.');
	});
});

describe('GCD / LCM', () => {
	it('computes greatest common divisors', () => {
		expect(f('GCD', 5, 2)).toBe(1);
		expect(f('GCD', 24, 36)).toBe(12);
		expect(f('GCD', 7, 1)).toBe(1);
		expect(f('GCD', 5, 0)).toBe(5);
		expect(f('GCD', 24.9, 36.1)).toBe(12);
		expect(f('GCD', '12', 18)).toBe(6);
		expect(f('GCD', ctx.ref('A1:A3'))).toBe(6);
		expect(f('GCD', ctx.ref('C1:C3'))).toBe(2);
		expect(f('GCD', arr([12, 30]), 42)).toBe(6);
	});

	it('computes least common multiples', () => {
		expect(f('LCM', 5, 2)).toBe(10);
		expect(f('LCM', 24, 36)).toBe(72);
		expect(f('LCM', 0, 5)).toBe(0);
		expect(f('LCM', ctx.ref('A1:A3'))).toBe(72);
		expect(f('LCM', 2 ** 30, 2 ** 30 - 1)).toBe(ERRORS.NUM);
	});

	it('rejects negatives, huge values, text in ranges and errors', () => {
		expect(f('GCD', -1, 2)).toBe(ERRORS.NUM);
		expect(f('LCM', -1, 2)).toBe(ERRORS.NUM);
		expect(f('GCD', 2 ** 53, 2)).toBe(ERRORS.NUM);
		expect(f('GCD', ctx.ref('B1:B2'))).toBe(ERRORS.VALUE);
		expect(f('GCD', 'abc')).toBe(ERRORS.VALUE);
		expect(f('LCM', ctx.ref('D1'), 3)).toBe(ERRORS.DIV0);
	});
});

describe('MULTINOMIAL', () => {
	it('computes multinomial coefficients', () => {
		expect(f('MULTINOMIAL', 2, 3, 4)).toBe(1260);
		expect(f('MULTINOMIAL', 1, 1)).toBe(2);
		expect(f('MULTINOMIAL', 0)).toBe(1);
		expect(f('MULTINOMIAL', 3.9, 1)).toBe(4);
		expect(f('MULTINOMIAL', arr([2, 3]), 4)).toBe(1260);
	});

	it('rejects negatives and text', () => {
		expect(f('MULTINOMIAL', -1, 2)).toBe(ERRORS.NUM);
		expect(f('MULTINOMIAL', 'x')).toBe(ERRORS.VALUE);
		expect(f('MULTINOMIAL', 1000, 1000)).toBe(ERRORS.NUM);
	});
});

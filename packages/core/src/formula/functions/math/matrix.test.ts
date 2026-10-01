import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { ArrayValue, FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 3, B1: 6, A2: 1, B2: 1, C1: 'x', C2: 2 });
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);

function expectMatrixClose(actual: unknown, expected: number[][]): void {
	const a = actual as ArrayValue;
	expect(a.kind).toBe('array');
	expect([a.rows, a.cols]).toEqual([expected.length, expected[0].length]);
	a.values.forEach((row, r) => row.forEach((v, c) => expect(v).toBeCloseTo(expected[r][c], 12)));
}

describe('MDETERM', () => {
	it('computes determinants', () => {
		expect(f('MDETERM', arr([1, 3, 8, 5], [1, 3, 6, 1], [1, 1, 1, 0], [7, 3, 10, 2]))).toBe(88);
		expect(f('MDETERM', arr([3, 6, 1], [1, 1, 0], [3, 10, 2]))).toBe(1);
		expect(f('MDETERM', arr([3, 6], [1, 1]))).toBe(-3);
		expect(f('MDETERM', arr([1, 2], [2, 4]))).toBe(0);
		expect(f('MDETERM', arr([0.5, 0.25], [0.1, 1]))).toBe(0.475);
		expect(f('MDETERM', 5)).toBe(5);
		expect(f('MDETERM', ctx.ref('A1:B2'))).toBe(-3);
	});

	it('rejects non-square or non-numeric input', () => {
		expect(f('MDETERM', arr([1, 3, 8], [4, 2, 5]))).toBe(ERRORS.VALUE);
		expect(f('MDETERM', arr([1, 'a'], [2, 3]))).toBe(ERRORS.VALUE);
		expect(f('MDETERM', arr([1, null], [2, 3]))).toBe(ERRORS.VALUE);
		expect(f('MDETERM', arr([1, ERRORS.NA], [2, 3]))).toBe(ERRORS.NA);
		expect(f('MDETERM', ctx.ref('B1:C2'))).toBe(ERRORS.VALUE);
	});
});

describe('MINVERSE', () => {
	it('inverts matrices', () => {
		expect(f('MINVERSE', arr([4, -1], [2, 0]))).toEqual(arr([0, 0.5], [-1, 2]));
		expectMatrixClose(f('MINVERSE', arr([1, 2, 1], [3, 4, -1], [0, 2, 0])), [
			[0.25, 0.25, -0.75],
			[0, 0, 0.5],
			[0.75, -0.25, -0.25],
		]);
		expect(f('MINVERSE', 4)).toEqual(arr([0.25]));
	});

	it('reports singular, non-square and non-numeric input', () => {
		expect(f('MINVERSE', arr([1, 2], [2, 4]))).toBe(ERRORS.NUM);
		expect(f('MINVERSE', arr([0, 0], [0, 0]))).toBe(ERRORS.NUM);
		expect(f('MINVERSE', arr([1, 2, 3]))).toBe(ERRORS.VALUE);
		expect(f('MINVERSE', arr([1, true], [2, 3]))).toBe(ERRORS.VALUE);
		expect(fn(MATH_FUNCTIONS, 'MINVERSE').returnsArray).toBe(true);
	});

	it('multiplies back to the identity', () => {
		const m = arr([2, -1, 0], [-1, 2, -1], [0, -1, 2]);
		const product = f('MMULT', m, f('MINVERSE', m) as ArrayValue);
		expectMatrixClose(product, [
			[1, 0, 0],
			[0, 1, 0],
			[0, 0, 1],
		]);
	});
});

describe('MMULT', () => {
	it('multiplies matrices', () => {
		expect(f('MMULT', arr([1, 3], [7, 2]), arr([2, 0], [0, 2]))).toEqual(arr([2, 6], [14, 4]));
		expect(f('MMULT', arr([1, 2]), arr([3], [4]))).toEqual(arr([11]));
		expect(f('MMULT', arr([1], [2]), arr([3, 4]))).toEqual(arr([3, 4], [6, 8]));
		expect(f('MMULT', ctx.ref('A1:B2'), arr([1], [1]))).toEqual(arr([9], [2]));
	});

	it('validates shapes and contents', () => {
		expect(f('MMULT', arr([1, 2]), arr([1, 2]))).toBe(ERRORS.VALUE);
		expect(f('MMULT', arr(['a']), arr([1]))).toBe(ERRORS.VALUE);
		expect(f('MMULT', arr([ERRORS.DIV0]), arr([1]))).toBe(ERRORS.DIV0);
	});
});

describe('MUNIT', () => {
	it('builds identity matrices', () => {
		expect(f('MUNIT', 3)).toEqual(arr([1, 0, 0], [0, 1, 0], [0, 0, 1]));
		expect(f('MUNIT', 1)).toEqual(arr([1]));
		expect(f('MUNIT', 2.9)).toEqual(arr([1, 0], [0, 1]));
	});

	it('rejects non-positive dimensions', () => {
		expect(f('MUNIT', 0)).toBe(ERRORS.VALUE);
		expect(f('MUNIT', -2)).toBe(ERRORS.VALUE);
		expect(f('MUNIT', 'a')).toBe(ERRORS.VALUE);
		expect(f('MUNIT', 100000)).toBe(ERRORS.VALUE);
	});
});

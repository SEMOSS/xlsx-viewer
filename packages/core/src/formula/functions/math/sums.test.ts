import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';
import { sumOf } from './sums';

const ctx = new MockContext({
	A1: 1,
	A2: 2,
	A3: 'text',
	A4: true,
	A5: '4',
	A6: 3,
	B1: ERRORS.NA,
	C1: 5,
	C2: 15,
	C3: 30,
});
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);
const missing: AstNode = { type: 'missing' };

describe('SUM', () => {
	it('adds direct arguments with coercion', () => {
		expect(f('SUM', 1, 2, 3)).toBe(6);
		expect(f('SUM', '3', 4)).toBe(7);
		expect(f('SUM', true, 1)).toBe(2);
		expect(f('SUM', 1, missing, 2)).toBe(3);
	});

	it('ignores text and booleans inside ranges and arrays', () => {
		expect(f('SUM', ctx.ref('A1:A6'))).toBe(6);
		expect(f('SUM', arr([1, '2', true]))).toBe(1);
		expect(f('SUM', ctx.ref('A1:A2'), ctx.ref('C1:C3'), 10)).toBe(63);
		expect(f('SUM', { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('C1')] })).toBe(6);
	});

	it('propagates errors and rejects non-numeric direct text', () => {
		expect(f('SUM', 'abc')).toBe(ERRORS.VALUE);
		expect(f('SUM', ctx.ref('A1:B1'))).toBe(ERRORS.NA);
		expect(f('SUM', 1, ERRORS.DIV0)).toBe(ERRORS.DIV0);
	});

	it('sums accurately', () => {
		expect(f('SUM', 0.1, 0.2, 0.3)).toBe(0.6);
		expect(sumOf([1e16, 1, -1e16])).toBe(1);
		expect(f('SUM', ctx.ref('Z1:Z10'))).toBe(0);
	});
});

describe('PRODUCT / SUMSQ', () => {
	it('multiplies', () => {
		expect(f('PRODUCT', ctx.ref('C1:C3'))).toBe(2250);
		expect(f('PRODUCT', ctx.ref('C1:C3'), 2)).toBe(4500);
		expect(f('PRODUCT', ctx.ref('A1:A6'))).toBe(6);
		expect(f('PRODUCT', ctx.ref('Z1:Z3'))).toBe(0);
		expect(f('PRODUCT', '2', 3)).toBe(6);
		expect(f('PRODUCT', ctx.ref('B1'))).toBe(ERRORS.NA);
	});

	it('sums squares', () => {
		expect(f('SUMSQ', 3, 4)).toBe(25);
		expect(f('SUMSQ', '3', 4)).toBe(25);
		expect(f('SUMSQ', ctx.ref('A1:A6'))).toBe(14);
		expect(f('SUMSQ', 'x')).toBe(ERRORS.VALUE);
	});
});

describe('SUMPRODUCT', () => {
	it('multiplies corresponding entries and sums', () => {
		expect(f('SUMPRODUCT', arr([3, 4], [8, 6], [1, 9]), arr([2, 7], [6, 7], [5, 3]))).toBe(156);
		expect(f('SUMPRODUCT', arr([1, 2, 3]))).toBe(6);
		expect(f('SUMPRODUCT', 3, 4)).toBe(12);
		expect(f('SUMPRODUCT', ctx.ref('C1:C3'), arr([1], [2], [3]))).toBe(125);
	});

	it('treats non-numeric entries as zero', () => {
		expect(f('SUMPRODUCT', ctx.ref('A1:A6'), ctx.ref('A1:A6'))).toBe(14);
		expect(f('SUMPRODUCT', arr([true, '2', 3]), arr([1, 1, 1]))).toBe(3);
	});

	it('requires equal dimensions and propagates errors', () => {
		expect(f('SUMPRODUCT', arr([1, 2]), arr([1, 2, 3]))).toBe(ERRORS.VALUE);
		expect(f('SUMPRODUCT', arr([1, 2]), arr([1], [2]))).toBe(ERRORS.VALUE);
		expect(f('SUMPRODUCT', arr([1, ERRORS.DIV0]), arr([1, 2]))).toBe(ERRORS.DIV0);
	});
});

describe('SUMX2MY2 / SUMX2PY2 / SUMXMY2', () => {
	const x = arr([2, 3, 9, 1, 8, 7, 5]);
	const y = arr([6, 5, 11, 7, 5, 4, 4]);

	it('computes the documented examples', () => {
		expect(f('SUMX2MY2', x, y)).toBe(-55);
		expect(f('SUMX2PY2', x, y)).toBe(521);
		expect(f('SUMXMY2', x, y)).toBe(79);
	});

	it('skips pairs with non-numeric values', () => {
		expect(f('SUMX2PY2', arr([1, 'a', 3]), arr([1, 2, null]))).toBe(2);
		expect(f('SUMXMY2', arr([1, 2]), arr([1], [4]))).toBe(4);
	});

	it('requires equal sizes and propagates errors', () => {
		expect(f('SUMX2MY2', arr([1, 2]), arr([1, 2, 3]))).toBe(ERRORS.NA);
		expect(f('SUMX2MY2', arr([1, ERRORS.NUM]), arr([1, 2]))).toBe(ERRORS.NUM);
	});
});

describe('SERIESSUM', () => {
	it('evaluates power series', () => {
		const cosine = f('SERIESSUM', Math.PI / 4, 0, 2, arr([1, -1 / 2, 1 / 24, -1 / 720]));
		expect(cosine).toBeCloseTo(0.707103, 6);
		expect(f('SERIESSUM', 2, 1, 1, arr([1, 1, 1]))).toBe(14);
		expect(f('SERIESSUM', 2, 0, 1, ctx.ref('C1:C2'))).toBe(35);
	});

	it('rejects non-numeric coefficients and propagates errors', () => {
		expect(f('SERIESSUM', 2, 1, 1, arr([1, 'a']))).toBe(ERRORS.VALUE);
		expect(f('SERIESSUM', 2, 1, 1, arr([1, ERRORS.NA]))).toBe(ERRORS.NA);
		expect(f('SERIESSUM', 'x', 1, 1, arr([1]))).toBe(ERRORS.VALUE);
		expect(f('SERIESSUM', 0, -1, 1, arr([1]))).toBe(ERRORS.DIV0);
	});
});

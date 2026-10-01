import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

function subtotalContext(): MockContext {
	const ctx = new MockContext({
		A1: 120,
		A2: 10,
		A3: 150,
		A4: 23,
		A5: 303, // a nested SUBTOTAL(9, A1:A4)
		B1: 'label',
		B2: true,
		B3: ERRORS.DIV0,
		B4: 4,
	});
	ctx.markSubtotal('A5');
	return ctx;
}

// AGGREGATE documentation sample.
function aggregateContext(): MockContext {
	return new MockContext({
		A1: ERRORS.DIV0,
		A2: 72,
		A3: 30,
		A4: ERRORS.NUM,
		A5: 31,
		A6: 96,
		A7: 32,
		A8: 81,
		A9: 33,
		A10: 53,
		B1: 82,
		B2: 65,
		B3: 95,
		B4: 63,
		B5: 53,
		B6: 71,
		B7: 55,
		B8: 83,
		B9: 100,
		B10: 91,
	});
}

describe('SUBTOTAL', () => {
	const ctx = subtotalContext();
	const f = (...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, 'SUBTOTAL'), args, ctx);

	it.each([
		[1, 75.75],
		[2, 4],
		[3, 4],
		[4, 150],
		[5, 10],
		[6, 4140000],
		[9, 303],
		[11, 3644.1875],
	])('SUBTOTAL(%i, A1:A4) = %d', (code, expected) => {
		expect(f(code, ctx.ref('A1:A4'))).toBe(expected);
	});

	it('computes sample statistics', () => {
		expect(f(7, ctx.ref('A1:A4'))).toBeCloseTo(69.70593, 5);
		expect(f(8, ctx.ref('A1:A4'))).toBeCloseTo(60.36711, 5);
		expect(f(10, ctx.ref('A1:A4'))).toBeCloseTo(4858.916667, 6);
	});

	it('skips nested SUBTOTAL / AGGREGATE cells', () => {
		expect(f(9, ctx.ref('A1:A5'))).toBe(303);
		expect(f(2, ctx.ref('A1:A5'))).toBe(4);
	});

	it('includes hidden rows for 1-11 and excludes them for 101-111', () => {
		const hidden = subtotalContext();
		hidden.hideRow(1);
		const g = (...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, 'SUBTOTAL'), args, hidden);
		expect(g(9, hidden.ref('A1:A4'))).toBe(303);
		expect(g(109, hidden.ref('A1:A4'))).toBe(293);
		expect(g(101, hidden.ref('A1:A4'))).toBeCloseTo(97.6667, 4);
		expect(g(102, hidden.ref('A1:A4'))).toBe(3);
	});

	it('handles mixed values and errors like the underlying functions', () => {
		expect(f(3, ctx.ref('B1:B4'))).toBe(4);
		expect(f(2, ctx.ref('B1:B4'))).toBe(1);
		expect(f(9, ctx.ref('B1:B4'))).toBe(ERRORS.DIV0);
		expect(f(9, ctx.ref('B4'))).toBe(4);
	});

	it('accepts several references and unions', () => {
		expect(f(9, ctx.ref('A1:A2'), ctx.ref('A4'))).toBe(153);
		expect(f(9, { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('A3')] })).toBe(270);
		expect(f(109, ctx.ref('Z1:Z5'))).toBe(0);
		expect(f(101, ctx.ref('Z1:Z5'))).toBe(ERRORS.DIV0);
	});

	it('rejects invalid function numbers, arrays and 3-D references', () => {
		expect(f(12, ctx.ref('A1:A4'))).toBe(ERRORS.VALUE);
		expect(f(0, ctx.ref('A1:A4'))).toBe(ERRORS.VALUE);
		expect(f(100, ctx.ref('A1:A4'))).toBe(ERRORS.VALUE);
		expect(f(9, arr([1, 2]))).toBe(ERRORS.VALUE);
		expect(f(9, 5)).toBe(ERRORS.VALUE);
		expect(f(9, { ...ctx.ref('A1:A4'), sheetEnd: 1 })).toBe(ERRORS.VALUE);
		expect(f('x', ctx.ref('A1:A4'))).toBe(ERRORS.VALUE);
		expect(f(ERRORS.NA, ctx.ref('A1:A4'))).toBe(ERRORS.NA);
	});

	it('lifts over an array of function numbers', () => {
		expect(f(arr([9, 4]), ctx.ref('A1:A4'))).toEqual(arr([303, 150]));
	});
});

describe('AGGREGATE', () => {
	const ctx = aggregateContext();
	const f = (...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, 'AGGREGATE'), args, ctx);

	it('computes the documented examples', () => {
		expect(f(4, 6, ctx.ref('A1:A10'))).toBe(96);
		expect(f(14, 6, ctx.ref('A1:A10'), 3)).toBe(72);
		expect(f(15, 6, ctx.ref('A1:A10'))).toBe(ERRORS.VALUE);
		expect(f(12, 6, ctx.ref('A1:A10'), ctx.ref('B1:B10'))).toBe(68);
		expect(f(4, 4, ctx.ref('A1:A10'))).toBe(ERRORS.DIV0);
	});

	it.each([
		[1, 53.5],
		[2, 8],
		[3, 8],
		[4, 96],
		[5, 30],
		[9, 428],
		[12, 43],
	])('function %i ignoring errors', (code, expected) => {
		expect(f(code, 6, ctx.ref('A1:A10'))).toBe(expected);
	});

	it('computes the remaining statistics', () => {
		expect(f(6, 6, ctx.ref('A2:A3'))).toBe(2160);
		expect(f(7, 6, ctx.ref('B1:B10'))).toBeCloseTo(16.77167, 5);
		expect(f(8, 6, ctx.ref('B1:B10'))).toBeCloseTo(15.911, 4);
		expect(f(10, 6, ctx.ref('B1:B10'))).toBeCloseTo(281.288889, 6);
		expect(f(11, 6, ctx.ref('B1:B10'))).toBeCloseTo(253.16, 10);
		expect(f(13, 6, ctx.ref('A1:B10'))).toBe(53);
		expect(f(13, 6, ctx.ref('A2:A3'))).toBe(ERRORS.NA);
	});

	it('counts errors as values for COUNTA unless they are ignored', () => {
		expect(f(3, 4, ctx.ref('A1:A10'))).toBe(10);
		expect(f(2, 4, ctx.ref('A1:A10'))).toBe(8);
		expect(f(3, 6, ctx.ref('A1:A10'))).toBe(8);
	});

	it('implements the array form (functions 14-19)', () => {
		const data = ctx.ref('B1:B10');
		expect(f(14, 0, data, 1)).toBe(100);
		expect(f(15, 0, data, 2)).toBe(55);
		expect(f(16, 0, data, 0.5)).toBe(76.5);
		expect(f(17, 0, data, 1)).toBe(63.5);
		expect(f(18, 0, data, 0.5)).toBe(76.5);
		expect(f(19, 0, data, 1)).toBe(61);
		expect(f(15, 6, arr([3, ERRORS.DIV0, 1, 2]), 1)).toBe(1);
		expect(f(14, 6, ctx.ref('A1:A10'), arr([1, 2]))).toEqual(arr([96, 81]));
	});

	it('validates array-form arguments', () => {
		const data = ctx.ref('B1:B10');
		expect(f(14, 0, data, 0)).toBe(ERRORS.NUM);
		expect(f(14, 0, data, 11)).toBe(ERRORS.NUM);
		expect(f(16, 0, data, 1.5)).toBe(ERRORS.NUM);
		expect(f(19, 0, data, 4)).toBe(ERRORS.NUM);
		expect(f(14, 0, data, 'x')).toBe(ERRORS.VALUE);
		expect(f(14, 0, ctx.ref('A1:A10'), 1)).toBe(ERRORS.DIV0);
		expect(f(14, 6, arr([1, 2]), 1, 2)).toBe(ERRORS.VALUE);
	});

	it('honours the hidden-row and nested-subtotal options', () => {
		const c = new MockContext({ A1: 1, A2: 2, A3: 4, A4: 7 });
		c.hideRow(1);
		c.markSubtotal('A4');
		const g = (...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, 'AGGREGATE'), args, c);
		const data = c.ref('A1:A4');
		expect(g(9, 0, data)).toBe(7);
		expect(g(9, 1, data)).toBe(5);
		expect(g(9, 4, data)).toBe(14);
		expect(g(9, 5, data)).toBe(12);
		expect(g(9, 3, data)).toBe(5);
		expect(g(14, 5, data, 1)).toBe(7);
		expect(g(14, 1, data, 1)).toBe(4);
	});

	it('rejects invalid function numbers, options and arrays in the reference form', () => {
		expect(f(20, 0, ctx.ref('B1:B10'))).toBe(ERRORS.VALUE);
		expect(f(0, 0, ctx.ref('B1:B10'))).toBe(ERRORS.VALUE);
		expect(f(9, 8, ctx.ref('B1:B10'))).toBe(ERRORS.VALUE);
		expect(f(9, -1, ctx.ref('B1:B10'))).toBe(ERRORS.VALUE);
		expect(f(9, 6, arr([1, 2]))).toBe(ERRORS.VALUE);
		expect(f(9, null, ctx.ref('B1:B3'))).toBe(242);
		expect(fn(MATH_FUNCTIONS, 'AGGREGATE').filePrefix).toBe('_xlfn.');
	});
});

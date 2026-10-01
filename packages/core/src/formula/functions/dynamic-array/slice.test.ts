import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { block } from '../lookup/test-utils/helpers';
import { CHOOSECOLS, CHOOSEROWS, DROP, EXPAND, TAKE } from './slice';

const MISSING = { type: 'missing' } as const;
const GRID = arr([1, 2, 3], [4, 5, 6], [7, 8, 9]);

describe('TAKE', () => {
	it('takes leading or trailing rows and columns', () => {
		expect(call(TAKE, [GRID, 2])).toEqual(arr([1, 2, 3], [4, 5, 6]));
		expect(call(TAKE, [GRID, -1])).toEqual(arr([7, 8, 9]));
		expect(call(TAKE, [GRID, MISSING, 2])).toEqual(arr([1, 2], [4, 5], [7, 8]));
		expect(call(TAKE, [GRID, 2, -1])).toEqual(arr([3], [6]));
		expect(call(TAKE, [GRID, -2, -2])).toEqual(arr([5, 6], [8, 9]));
	});

	it('keeps everything for counts beyond the array', () => {
		expect(call(TAKE, [GRID, 10])).toEqual(GRID);
		expect(call(TAKE, [GRID, -10, 10])).toEqual(GRID);
		expect(call(TAKE, [GRID, null, null])).toEqual(GRID);
	});

	it('truncates counts and rejects zero', () => {
		expect(call(TAKE, [GRID, 1.9])).toEqual(arr([1, 2, 3]));
		expect(call(TAKE, [GRID, -1.9])).toEqual(arr([7, 8, 9]));
		expect(call(TAKE, [GRID, 0])).toBe(ERRORS.CALC);
		expect(call(TAKE, [GRID, 1, 0])).toBe(ERRORS.CALC);
		expect(call(TAKE, [GRID, 0.5])).toBe(ERRORS.CALC);
	});

	it('propagates errors and reads references', () => {
		expect(call(TAKE, [GRID, 'x'])).toBe(ERRORS.VALUE);
		expect(call(TAKE, [GRID, ERRORS.NA])).toBe(ERRORS.NA);
		const ctx = new MockContext(
			block([
				['a', 'b'],
				['c', null],
			]),
		);
		expect(call(TAKE, [ctx.ref('A1:B2'), -1], ctx)).toEqual(arr(['c', null]));
		expect(call(TAKE, [5, 1])).toEqual(arr([5]));
	});
});

describe('DROP', () => {
	it('drops leading or trailing rows and columns', () => {
		expect(call(DROP, [GRID, 1])).toEqual(arr([4, 5, 6], [7, 8, 9]));
		expect(call(DROP, [GRID, -2])).toEqual(arr([1, 2, 3]));
		expect(call(DROP, [GRID, MISSING, 1])).toEqual(arr([2, 3], [5, 6], [8, 9]));
		expect(call(DROP, [GRID, 1, -1])).toEqual(arr([4, 5], [7, 8]));
		expect(call(DROP, [GRID, 0])).toEqual(GRID);
		expect(call(DROP, [GRID, 0, 0])).toEqual(GRID);
	});

	it('is #CALC! when everything is dropped', () => {
		expect(call(DROP, [GRID, 3])).toBe(ERRORS.CALC);
		expect(call(DROP, [GRID, -5])).toBe(ERRORS.CALC);
		expect(call(DROP, [GRID, 1, 3])).toBe(ERRORS.CALC);
		expect(call(DROP, [GRID, 2.7])).toEqual(arr([7, 8, 9]));
	});
});

describe('CHOOSECOLS / CHOOSEROWS', () => {
	it('picks columns in the given order', () => {
		expect(call(CHOOSECOLS, [GRID, 3, 1])).toEqual(arr([3, 1], [6, 4], [9, 7]));
		expect(call(CHOOSECOLS, [GRID, -1])).toEqual(arr([3], [6], [9]));
		expect(call(CHOOSECOLS, [GRID, arr([1, 2]), 2])).toEqual(arr([1, 2, 2], [4, 5, 5], [7, 8, 8]));
		expect(call(CHOOSECOLS, [GRID, '2', 2.9])).toEqual(arr([2, 2], [5, 5], [8, 8]));
	});

	it('picks rows in the given order', () => {
		expect(call(CHOOSEROWS, [GRID, 3, 1])).toEqual(arr([7, 8, 9], [1, 2, 3]));
		expect(call(CHOOSEROWS, [GRID, -2])).toEqual(arr([4, 5, 6]));
		expect(call(CHOOSEROWS, [GRID, arr([1], [1])])).toEqual(arr([1, 2, 3], [1, 2, 3]));
	});

	it('rejects zero and out-of-range positions', () => {
		expect(call(CHOOSECOLS, [GRID, 0])).toBe(ERRORS.VALUE);
		expect(call(CHOOSECOLS, [GRID, 4])).toBe(ERRORS.VALUE);
		expect(call(CHOOSECOLS, [GRID, -4])).toBe(ERRORS.VALUE);
		expect(call(CHOOSEROWS, [GRID, 1, 9])).toBe(ERRORS.VALUE);
		expect(call(CHOOSEROWS, [GRID, 'x'])).toBe(ERRORS.VALUE);
		expect(call(CHOOSEROWS, [GRID, ERRORS.DIV0])).toBe(ERRORS.DIV0);
		expect(call(CHOOSEROWS, [GRID, MISSING])).toBe(ERRORS.VALUE);
	});

	it('reads references', () => {
		const ctx = new MockContext(
			block([
				[1, 2],
				[3, 4],
			]),
		);
		expect(call(CHOOSECOLS, [ctx.ref('A1:B2'), 2], ctx)).toEqual(arr([2], [4]));
		expect(call(CHOOSEROWS, [ctx.ref('A1:B2'), ctx.ref('B1')], ctx)).toEqual(arr([3, 4]));
	});
});

describe('EXPAND', () => {
	it('pads with #N/A by default', () => {
		expect(call(EXPAND, [arr([1, 2]), 2])).toEqual(arr([1, 2], [ERRORS.NA, ERRORS.NA]));
		expect(call(EXPAND, [arr([1], [2]), MISSING, 2])).toEqual(arr([1, ERRORS.NA], [2, ERRORS.NA]));
	});

	it('pads with pad_with, errors included', () => {
		expect(call(EXPAND, [arr([1]), 2, 2, 0])).toEqual(arr([1, 0], [0, 0]));
		expect(call(EXPAND, [arr([1]), 1, 3, ''])).toEqual(arr([1, '', '']));
		expect(call(EXPAND, [arr([1]), 1, 2, ERRORS.DIV0])).toEqual(arr([1, ERRORS.DIV0]));
		expect(call(EXPAND, [arr([1]), 2, 1, MISSING])).toEqual(arr([1], [ERRORS.NA]));
	});

	it('keeps the size when omitted or equal', () => {
		expect(call(EXPAND, [GRID, 3, 3])).toEqual(GRID);
		expect(call(EXPAND, [GRID, MISSING, MISSING])).toEqual(GRID);
	});

	it('rejects shrinking, bad sizes and huge results', () => {
		expect(call(EXPAND, [GRID, 2])).toBe(ERRORS.VALUE);
		expect(call(EXPAND, [GRID, 3, 1])).toBe(ERRORS.VALUE);
		expect(call(EXPAND, [GRID, -1])).toBe(ERRORS.VALUE);
		expect(call(EXPAND, [GRID, 'x'])).toBe(ERRORS.VALUE);
		expect(call(EXPAND, [GRID, ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(EXPAND, [GRID, 1_000_000, 100])).toBe(ERRORS.NUM);
	});
});

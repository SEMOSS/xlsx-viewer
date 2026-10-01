import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { RANDARRAY, SEQUENCE } from './sequence';

const MISSING = { type: 'missing' } as const;

describe('SEQUENCE', () => {
	it('counts down rows from 1 by default', () => {
		expect(call(SEQUENCE, [3])).toEqual(arr([1], [2], [3]));
		expect(call(SEQUENCE, [1])).toEqual(arr([1]));
	});

	it('fills row by row with start and step', () => {
		expect(call(SEQUENCE, [2, 3])).toEqual(arr([1, 2, 3], [4, 5, 6]));
		expect(call(SEQUENCE, [2, 2, 10, 5])).toEqual(arr([10, 15], [20, 25]));
		expect(call(SEQUENCE, [1, 4, 0, -0.5])).toEqual(arr([0, -0.5, -1, -1.5]));
		expect(call(SEQUENCE, [3, 1, 5, 0])).toEqual(arr([5], [5], [5]));
	});

	it('defaults omitted arguments to 1', () => {
		expect(call(SEQUENCE, [3, MISSING, 0])).toEqual(arr([0], [1], [2]));
		expect(call(SEQUENCE, [2, null, null, 10])).toEqual(arr([1], [11]));
		expect(call(SEQUENCE, [MISSING, 3])).toEqual(arr([1, 2, 3]));
	});

	it('truncates sizes and coerces arguments', () => {
		expect(call(SEQUENCE, [2.9])).toEqual(arr([1], [2]));
		expect(call(SEQUENCE, ['2', true])).toEqual(arr([1], [2]));
	});

	it('rejects empty, negative and oversized results', () => {
		expect(call(SEQUENCE, [0])).toBe(ERRORS.CALC);
		expect(call(SEQUENCE, [3, 0])).toBe(ERRORS.CALC);
		expect(call(SEQUENCE, [0.5])).toBe(ERRORS.CALC);
		expect(call(SEQUENCE, [-1])).toBe(ERRORS.VALUE);
		expect(call(SEQUENCE, [2, -3])).toBe(ERRORS.VALUE);
		expect(call(SEQUENCE, [100_000, 1000])).toBe(ERRORS.NUM);
		expect(call(SEQUENCE, ['x'])).toBe(ERRORS.VALUE);
		expect(call(SEQUENCE, [2, 2, 'start'])).toBe(ERRORS.VALUE);
		expect(call(SEQUENCE, [ERRORS.DIV0])).toBe(ERRORS.DIV0);
	});

	it('handles a 1-million-cell result', () => {
		const out = call(SEQUENCE, [1000, 1000]);
		expect(out).toMatchObject({ rows: 1000, cols: 1000 });
	});
});

describe('RANDARRAY', () => {
	function ctxWith(...values: number[]): MockContext {
		const ctx = new MockContext();
		ctx.randomValues = values;
		return ctx;
	}

	it('returns one random number in [0, 1) by default', () => {
		expect(call(RANDARRAY, [], ctxWith(0.25))).toEqual(arr([0.25]));
	});

	it('fills rows and columns in row-major order', () => {
		expect(call(RANDARRAY, [2, 2], ctxWith(0.1, 0.2, 0.3, 0.4))).toEqual(
			arr([0.1, 0.2], [0.3, 0.4]),
		);
	});

	it('scales into [min, max)', () => {
		expect(call(RANDARRAY, [1, 2, 10, 20], ctxWith(0, 0.5))).toEqual(arr([10, 15]));
		expect(call(RANDARRAY, [1, 1, -1, 1], ctxWith(0.75))).toEqual(arr([0.5]));
		expect(call(RANDARRAY, [1, 1, 5, 5], ctxWith(0.9))).toEqual(arr([5]));
	});

	it('produces whole numbers in [min, max] when integer is TRUE', () => {
		expect(call(RANDARRAY, [1, 3, 1, 6, true], ctxWith(0, 0.5, 0.9999))).toEqual(arr([1, 4, 6]));
		expect(call(RANDARRAY, [1, 1, -3, -1, true], ctxWith(0.5))).toEqual(arr([-2]));
	});

	it('validates its arguments', () => {
		expect(call(RANDARRAY, [1, 1, 5, 1])).toBe(ERRORS.VALUE);
		expect(call(RANDARRAY, [1, 1, 1.5, 3, true])).toBe(ERRORS.VALUE);
		expect(call(RANDARRAY, [1, 1, 1, 3.5, true])).toBe(ERRORS.VALUE);
		expect(call(RANDARRAY, [0])).toBe(ERRORS.CALC);
		expect(call(RANDARRAY, [-2])).toBe(ERRORS.VALUE);
		expect(call(RANDARRAY, [1, 1, 0, 1, 'perhaps'])).toBe(ERRORS.VALUE);
		expect(call(RANDARRAY, [ERRORS.NA])).toBe(ERRORS.NA);
	});

	it('defaults omitted arguments', () => {
		expect(call(RANDARRAY, [MISSING, 2, MISSING, 10], ctxWith(0.5, 0.1))).toEqual(arr([5, 1]));
		expect(RANDARRAY.volatile).toBe(true);
	});
});

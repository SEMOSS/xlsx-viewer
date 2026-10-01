import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { rowsOf, stat } from '../test-utils/expect';

describe('MAX / MIN / MAXA / MINA', () => {
	it('matches Excel', () => {
		expect(stat('MAX', [arr([10, 7, 9, 27, 2])])).toBe(27);
		expect(stat('MAX', [arr([10, 7, 9, 27, 2]), 30])).toBe(30);
		expect(stat('MIN', [arr([10, 7, 9, 27, 2])])).toBe(2);
		expect(stat('MIN', [arr([10, 7, 9, 27, 2]), 0])).toBe(0);
	});

	it('returns 0 without numbers and ignores text in references', () => {
		const ctx = new MockContext({ A1: 'x', A2: true, A3: -5 });
		expect(stat('MAX', [ctx.ref('A1:A2')], ctx)).toBe(0);
		expect(stat('MIN', [ctx.ref('A1:A3')], ctx)).toBe(-5);
		expect(stat('MAX', [ctx.ref('A1:A3')], ctx)).toBe(-5);
		expect(stat('MAX', ['7', true])).toBe(7);
	});

	it('counts logical values and text in references for MAXA / MINA', () => {
		const ctx = new MockContext({
			A1: 0,
			A2: 0.2,
			A3: 0.5,
			A4: 0.4,
			A5: true,
			B1: false,
			B2: 0.2,
			B3: 'x',
			B4: 0.8,
		});
		expect(stat('MAXA', [ctx.ref('A1:A5')], ctx)).toBe(1);
		expect(stat('MINA', [ctx.ref('B1:B4')], ctx)).toBe(0);
		expect(stat('MAXA', [ctx.ref('B3')], ctx)).toBe(0);
	});

	it('propagates errors', () => {
		expect(stat('MAX', [arr([1, ERRORS.NA])])).toBe(ERRORS.NA);
		expect(stat('MIN', [1, 'abc'])).toBe(ERRORS.VALUE);
	});
});

describe('LARGE / SMALL', () => {
	const data = arr([3, 5, 3, 5, 4], [4, 2, 4, 6, 7]);

	it('matches Excel', () => {
		expect(stat('LARGE', [data, 3])).toBe(5);
		expect(stat('LARGE', [data, 7])).toBe(4);
		expect(stat('SMALL', [arr([3, 4, 5, 2, 3, 4, 6, 4, 7]), 4])).toBe(4);
		expect(stat('SMALL', [arr([1, 4, 8, 3, 7, 12, 54, 8, 23]), 2])).toBe(3);
	});

	it('rounds k up and validates it', () => {
		expect(stat('LARGE', [data, 1.2])).toBe(stat('LARGE', [data, 2]));
		expect(stat('SMALL', [data, 10])).toBe(7);
		expect(stat('SMALL', [data, 11])).toBe(ERRORS.NUM);
		expect(stat('LARGE', [data, 0])).toBe(ERRORS.NUM);
		expect(stat('LARGE', [arr(['a']), 1])).toBe(ERRORS.NUM);
	});

	it('is lifted over an array of k values', () => {
		expect(rowsOf(stat('LARGE', [data, arr([1, 2, 3])]))).toEqual([[7, 6, 5]]);
	});

	it('reads references, skipping non-numeric cells', () => {
		const ctx = new MockContext({ A1: 5, A2: 'x', A3: 1, A4: 9 });
		expect(stat('SMALL', [ctx.ref('A1:A4'), 1], ctx)).toBe(1);
		expect(stat('LARGE', [ctx.ref('A1:A4'), 3], ctx)).toBe(1);
	});
});

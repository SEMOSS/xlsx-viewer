import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { column, row } from './test-utils/helpers';
import { XMATCH } from './xlookup';

const MISSING = { type: 'missing' } as const;
const FRUIT = arr(['apple'], ['Banana'], ['cherry'], ['banana'], ['date']);

describe('XMATCH', () => {
	it('returns 1-based positions of exact matches', () => {
		expect(call(XMATCH, ['cherry', FRUIT])).toBe(3);
		expect(call(XMATCH, ['BANANA', FRUIT])).toBe(2);
		expect(call(XMATCH, ['banana', FRUIT, 0, -1])).toBe(4);
		expect(call(XMATCH, ['fig', FRUIT])).toBe(ERRORS.NA);
		expect(call(XMATCH, [3, arr([1, 2, 3])])).toBe(3);
		expect(call(XMATCH, ['3', arr([1, 2, 3])])).toBe(ERRORS.NA);
	});

	it('finds the next smaller or larger item', () => {
		const sizes = arr([5], [25], [10], [50], [100]);
		expect(call(XMATCH, [26, sizes, -1])).toBe(2);
		expect(call(XMATCH, [26, sizes, 1])).toBe(4);
		expect(call(XMATCH, [1, sizes, -1])).toBe(ERRORS.NA);
		expect(call(XMATCH, [1000, sizes, 1])).toBe(ERRORS.NA);
		expect(call(XMATCH, [10, sizes, 1])).toBe(3);
	});

	it('matches wildcards and regular expressions', () => {
		expect(call(XMATCH, ['b*', FRUIT, 2])).toBe(2);
		expect(call(XMATCH, ['b*', FRUIT, 2, -1])).toBe(4);
		expect(call(XMATCH, ['?ate', FRUIT, 2])).toBe(5);
		expect(call(XMATCH, ['b*', FRUIT, 0])).toBe(ERRORS.NA);
		expect(call(XMATCH, ['^b', FRUIT, 3])).toBe(4);
		expect(call(XMATCH, ['an{2}', FRUIT, 3])).toBe(ERRORS.NA);
		expect(call(XMATCH, ['(an){2}', FRUIT, 3])).toBe(2);
	});

	it('runs binary searches', () => {
		const sorted = arr([1, 3, 5, 7, 9, 11]);
		expect(call(XMATCH, [7, sorted, 0, 2])).toBe(4);
		expect(call(XMATCH, [8, sorted, 0, 2])).toBe(ERRORS.NA);
		expect(call(XMATCH, [8, sorted, -1, 2])).toBe(4);
		expect(call(XMATCH, [8, sorted, 1, 2])).toBe(5);
		const reversed = arr([11, 9, 7, 5, 3, 1]);
		expect(call(XMATCH, [7, reversed, 0, -2])).toBe(3);
		expect(call(XMATCH, [8, reversed, -1, -2])).toBe(3);
		expect(call(XMATCH, [8, reversed, 1, -2])).toBe(2);
		expect(call(XMATCH, ['b*', FRUIT, 2, 2])).toBe(ERRORS.VALUE);
	});

	it('searches references in either orientation', () => {
		const ctx = new MockContext({ ...column('A', [10, 20, 30]), ...row(5, ['x', 'y', 'z']) });
		expect(call(XMATCH, [20, ctx.ref('A1:A3')], ctx)).toBe(2);
		expect(call(XMATCH, ['z', ctx.ref('A5:C5')], ctx)).toBe(3);
		expect(call(XMATCH, [30, ctx.ref('A:A')], ctx)).toBe(3);
		expect(call(XMATCH, [null, ctx.ref('A:A')], ctx)).toBe(4);
		expect(call(XMATCH, [null, ctx.ref('A1:A3')], ctx)).toBe(ERRORS.NA);
	});

	it('validates arguments', () => {
		expect(call(XMATCH, [1, arr([1, 2], [3, 4])])).toBe(ERRORS.VALUE);
		expect(call(XMATCH, [1, arr([1]), 5])).toBe(ERRORS.VALUE);
		expect(call(XMATCH, [1, arr([1]), 0, 0])).toBe(ERRORS.VALUE);
		expect(call(XMATCH, [ERRORS.DIV0, arr([1])])).toBe(ERRORS.DIV0);
		expect(call(XMATCH, [1, ERRORS.REF])).toBe(ERRORS.REF);
		expect(call(XMATCH, [1, arr([1, 1]), MISSING, -1])).toBe(2);
	});

	it('lifts over an array of lookup values', () => {
		expect(call(XMATCH, [arr(['date', 'apple', 'kiwi']), FRUIT])).toEqual(arr([5, 1, ERRORS.NA]));
	});
});

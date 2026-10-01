import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { MATCH } from './match';
import { column, CountingContext, row } from './test-utils/helpers';

const ASC = arr([10, 20, 30, 40]);
const DESC = arr([40, 30, 20, 10]);

describe('MATCH exact (match_type 0)', () => {
	it('finds numbers, text and logicals', () => {
		expect(call(MATCH, [30, ASC, 0])).toBe(3);
		expect(call(MATCH, [10, ASC, 0])).toBe(1);
		expect(call(MATCH, ['B', arr(['a', 'b', 'c']), 0])).toBe(2);
		expect(call(MATCH, [true, arr([1, false, true]), 0])).toBe(3);
		expect(call(MATCH, [false, arr([0, false]), 0])).toBe(2);
	});

	it('is #N/A when absent and never coerces types', () => {
		expect(call(MATCH, [25, ASC, 0])).toBe(ERRORS.NA);
		expect(call(MATCH, ['30', ASC, 0])).toBe(ERRORS.NA);
		expect(call(MATCH, [1, arr([true]), 0])).toBe(ERRORS.NA);
		expect(call(MATCH, [1, arr(['1']), 0])).toBe(ERRORS.NA);
	});

	it('supports wildcards', () => {
		const names = arr(['apple', 'banana', 'cherry', 'a*c']);
		expect(call(MATCH, ['b*', names, 0])).toBe(2);
		expect(call(MATCH, ['*rr*', names, 0])).toBe(3);
		expect(call(MATCH, ['?????', names, 0])).toBe(1);
		expect(call(MATCH, ['??????', names, 0])).toBe(2);
		expect(call(MATCH, ['a~*c', names, 0])).toBe(4);
		expect(call(MATCH, ['*', arr([1, 'x']), 0])).toBe(2);
		expect(call(MATCH, ['x*', names, 0])).toBe(ERRORS.NA);
	});

	it('returns the first of duplicates', () => {
		expect(call(MATCH, [2, arr([1, 2, 2, 3]), 0])).toBe(2);
		expect(call(MATCH, ['A', arr(['b', 'a', 'A']), 0])).toBe(2);
	});

	it('treats an empty match_type argument as 0', () => {
		expect(call(MATCH, [25, ASC, null])).toBe(ERRORS.NA);
		expect(call(MATCH, [25, ASC, { type: 'missing' }])).toBe(ERRORS.NA);
		expect(call(MATCH, [25, ASC])).toBe(2);
	});

	it('works down a column reference with blanks and errors', () => {
		const ctx = new MockContext(column('A', ['x', null, ERRORS.NA, 'y', '']));
		expect(call(MATCH, ['y', ctx.ref('A1:A5'), 0], ctx)).toBe(4);
		expect(call(MATCH, ['', ctx.ref('A1:A5'), 0], ctx)).toBe(5);
		expect(call(MATCH, [null, ctx.ref('A1:A5'), 0], ctx)).toBe(5);
		expect(call(MATCH, [0, ctx.ref('A1:A5'), 0], ctx)).toBe(ERRORS.NA);
	});
});

describe('MATCH approximate (match_type 1 / -1)', () => {
	it('defaults to the largest value <= lookup_value', () => {
		expect(call(MATCH, [25, ASC])).toBe(2);
		expect(call(MATCH, [25, ASC, 1])).toBe(2);
		expect(call(MATCH, [40, ASC, 1])).toBe(4);
		expect(call(MATCH, [45, ASC, 1])).toBe(4);
		expect(call(MATCH, [10, ASC, 1])).toBe(1);
		expect(call(MATCH, [5, ASC, 1])).toBe(ERRORS.NA);
	});

	it('returns the last of duplicates', () => {
		expect(call(MATCH, [2, arr([1, 2, 2, 2, 3]), 1])).toBe(4);
		expect(call(MATCH, ['b', arr(['a', 'b', 'B', 'c']), 1])).toBe(3);
	});

	it('treats any positive or negative match_type as 1 or -1', () => {
		expect(call(MATCH, [25, ASC, 2])).toBe(2);
		expect(call(MATCH, [25, ASC, 0.5])).toBe(2);
		expect(call(MATCH, [25, DESC, -5])).toBe(2);
		expect(call(MATCH, [25, ASC, true])).toBe(2);
		expect(call(MATCH, [25, ASC, '1'])).toBe(2);
	});

	it('finds the smallest value >= lookup_value in descending data', () => {
		expect(call(MATCH, [25, DESC, -1])).toBe(2);
		expect(call(MATCH, [30, DESC, -1])).toBe(2);
		expect(call(MATCH, [5, DESC, -1])).toBe(4);
		expect(call(MATCH, [40, DESC, -1])).toBe(1);
		expect(call(MATCH, [45, DESC, -1])).toBe(ERRORS.NA);
		expect(call(MATCH, [3, arr([5, 3, 3, 1]), -1])).toBe(3);
	});

	it('compares text alphabetically and case-insensitively', () => {
		const letters = arr(['A', 'c', 'E', 'g']);
		expect(call(MATCH, ['d', letters, 1])).toBe(2);
		expect(call(MATCH, ['G', letters, 1])).toBe(4);
		expect(call(MATCH, ['0', letters, 1])).toBe(ERRORS.NA);
		expect(call(MATCH, ['d', arr(['g', 'E', 'c', 'A']), -1])).toBe(2);
	});

	it('finds the last number with the 9.99E+307 idiom', () => {
		const ctx = new MockContext(column('A', ['Header', 1, 5, 'note', null, 7, 'x', 3, 'y']));
		expect(call(MATCH, [9.99e307, ctx.ref('A:A')], ctx)).toBe(8);
		expect(call(MATCH, [9.99e307, ctx.ref('A1:A100')], ctx)).toBe(8);
	});

	it('finds the last text with the REPT("z", 255) idiom', () => {
		const ctx = new MockContext(column('A', ['Header', 1, 5, 'note', null, 7, 'x', 3, 'y', 4]));
		expect(call(MATCH, ['z'.repeat(255), ctx.ref('A:A')], ctx)).toBe(9);
	});

	it('finds the last logical among other values', () => {
		expect(call(MATCH, [true, arr([1, true, 'a', false, 2])])).toBe(4);
	});

	it('is #N/A when no value of the lookup type exists', () => {
		expect(call(MATCH, [5, arr(['a', 'b'])])).toBe(ERRORS.NA);
		expect(call(MATCH, ['a', arr([1, 2])])).toBe(ERRORS.NA);
		expect(call(MATCH, [5, arr([null, ERRORS.NA])])).toBe(ERRORS.NA);
	});
});

describe('MATCH arguments', () => {
	it('requires a single row or column', () => {
		expect(call(MATCH, [1, arr([1, 2], [3, 4]), 0])).toBe(ERRORS.NA);
		expect(call(MATCH, [1, arr([1]), 0])).toBe(1);
		expect(call(MATCH, [7, 7, 0])).toBe(1);
	});

	it('searches a horizontal reference', () => {
		const ctx = new MockContext(row(2, ['a', 'b', 'c'], 1));
		expect(call(MATCH, ['c', ctx.ref('B2:D2'), 0], ctx)).toBe(3);
		expect(call(MATCH, ['bb', ctx.ref('B2:D2'), 1], ctx)).toBe(2);
	});

	it('propagates errors', () => {
		expect(call(MATCH, [ERRORS.DIV0, ASC, 0])).toBe(ERRORS.DIV0);
		expect(call(MATCH, [1, ERRORS.REF, 0])).toBe(ERRORS.REF);
		expect(call(MATCH, [1, ASC, ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(MATCH, [1, ASC, 'abc'])).toBe(ERRORS.VALUE);
	});

	it('rejects lookup text longer than 255 characters', () => {
		expect(call(MATCH, ['a'.repeat(256), arr(['a']), 0])).toBe(ERRORS.VALUE);
	});

	it('lifts over an array of lookup values', () => {
		expect(call(MATCH, [arr([20, 40, 5]), ASC, 0])).toEqual(arr([2, 4, ERRORS.NA]));
		expect(call(MATCH, [arr([15], [99]), ASC])).toEqual(arr([1], [4]));
	});

	it('stays fast on 100k sorted rows', () => {
		const cells: Record<string, number> = {};
		for (let r = 1; r <= 100_000; r++) {
			cells[`A${r}`] = r * 10;
		}
		const ctx = new CountingContext(cells);
		const range = ctx.ref('A1:A100000');
		expect(call(MATCH, [123_455, range, 1], ctx)).toBe(12_345);
		expect(ctx.reads).toBeLessThanOrEqual(20);
		ctx.reads = 0;
		expect(call(MATCH, [2_000_000, range, -1], ctx)).toBe(ERRORS.NA);
		expect(ctx.reads).toBeLessThanOrEqual(20);
		ctx.reads = 0;
		expect(call(MATCH, [1_000_000, ctx.ref('A:A'), 0], ctx)).toBe(100_000);
		expect(ctx.reads).toBe(100_000);
	});
});

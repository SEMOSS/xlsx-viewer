import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { ArrayValue, Scalar } from '../../types';
import { block } from '../lookup/test-utils/helpers';
import { SORT, SORTBY } from './sort';
import { compareForSort, sortPermutation } from './sort-order';

const MISSING = { type: 'missing' } as const;
const PEOPLE = arr(
	['Cy', 'East', 30],
	['ann', 'West', 20],
	['Bob', 'East', 20],
	['Di', 'West', 10],
);

function col(a: unknown, c = 0): Scalar[] {
	return (a as ArrayValue).values.map((r) => r[c]);
}

describe('compareForSort', () => {
	it('orders numbers, text, logicals, errors, then blanks', () => {
		const values: Scalar[] = [null, ERRORS.NA, true, false, 'b', 'A', 10, -1];
		const sorted = values.slice().sort((a, b) => compareForSort(a, b, false));
		expect(sorted).toEqual([-1, 10, 'A', 'b', false, true, ERRORS.NA, null]);
		const desc = values.slice().sort((a, b) => compareForSort(a, b, true));
		expect(desc).toEqual([ERRORS.NA, true, false, 'b', 'A', 10, -1, null]);
	});

	it('compares text case-insensitively and errors as equal', () => {
		expect(compareForSort('abc', 'ABC', false)).toBe(0);
		expect(compareForSort(ERRORS.NA, ERRORS.DIV0, false)).toBe(0);
		expect(compareForSort(null, null, true)).toBe(0);
		expect(compareForSort('', 'a', false)).toBeLessThan(0);
	});

	it('builds stable multi-key permutations', () => {
		expect(sortPermutation(4, [[2, 1, 2, 1]], [false])).toEqual([1, 3, 0, 2]);
		expect(
			sortPermutation(
				4,
				[
					[2, 1, 2, 1],
					['a', 'b', 'c', 'd'],
				],
				[false, true],
			),
		).toEqual([3, 1, 2, 0]);
	});
});

describe('SORT', () => {
	it('sorts rows by the first column ascending by default', () => {
		expect(call(SORT, [arr([3], [1], [2])])).toEqual(arr([1], [2], [3]));
		expect(col(call(SORT, [PEOPLE]))).toEqual(['ann', 'Bob', 'Cy', 'Di']);
	});

	it('sorts by another column, descending, and by several keys', () => {
		expect(col(call(SORT, [PEOPLE, 3]), 2)).toEqual([10, 20, 20, 30]);
		expect(col(call(SORT, [PEOPLE, 3]))).toEqual(['Di', 'ann', 'Bob', 'Cy']);
		expect(col(call(SORT, [PEOPLE, 3, -1]))).toEqual(['Cy', 'ann', 'Bob', 'Di']);
		expect(col(call(SORT, [PEOPLE, arr([2, 3]), arr([1, -1])]))).toEqual([
			'Cy',
			'Bob',
			'ann',
			'Di',
		]);
		expect(col(call(SORT, [PEOPLE, arr([2, 3]), -1]))).toEqual(['ann', 'Di', 'Cy', 'Bob']);
		expect(col(call(SORT, [PEOPLE, arr([2], [3]), arr([-1], [1])]))).toEqual([
			'Di',
			'ann',
			'Bob',
			'Cy',
		]);
	});

	it('sorts columns with by_col', () => {
		expect(call(SORT, [arr([3, 1, 2], ['c', 'a', 'b']), 1, 1, true])).toEqual(
			arr([1, 2, 3], ['a', 'b', 'c']),
		);
		expect(call(SORT, [arr([3, 1, 2], ['x', 'z', 'y']), 2, -1, true])).toEqual(
			arr([1, 2, 3], ['z', 'y', 'x']),
		);
	});

	it('uses Excel order for mixed values with blanks last', () => {
		const mixed = arr(['b'], [2], [true], [ERRORS.NA], [null], [1], ['A'], [false]);
		expect(col(call(SORT, [mixed]))).toEqual([1, 2, 'A', 'b', false, true, ERRORS.NA, null]);
		expect(col(call(SORT, [mixed, 1, -1]))).toEqual([ERRORS.NA, true, false, 'b', 'A', 2, 1, null]);
	});

	it('is stable for equal keys', () => {
		const a = arr(['b', 1], ['B', 2], ['a', 3], ['b', 4]);
		expect(col(call(SORT, [a]), 1)).toEqual([3, 1, 2, 4]);
		expect(col(call(SORT, [a, 1, -1]), 1)).toEqual([1, 2, 4, 3]);
	});

	it('defaults omitted arguments', () => {
		expect(col(call(SORT, [PEOPLE, MISSING, -1]))).toEqual(['Di', 'Cy', 'Bob', 'ann']);
		expect(col(call(SORT, [PEOPLE, null, null, null]))).toEqual(['ann', 'Bob', 'Cy', 'Di']);
		expect(col(call(SORT, [PEOPLE, 1, arr([-1])]))).toEqual(['Di', 'Cy', 'Bob', 'ann']);
		// A blank value (unlike an omitted argument) reads as 0.
		expect(call(SORT, [PEOPLE, arr([null])])).toBe(ERRORS.VALUE);
	});

	it('validates sort_index and sort_order', () => {
		expect(call(SORT, [PEOPLE, 0])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, 4])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, 'x'])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, arr([1, 2], [3, 1])])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, 1, 2])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, 1, 0])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, arr([1, 2]), arr([1, -1, 1])])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, 1, 1, 'maybe'])).toBe(ERRORS.VALUE);
		expect(call(SORT, [PEOPLE, ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(SORT, [PEOPLE, 1, arr([ERRORS.DIV0])])).toBe(ERRORS.DIV0);
		expect(call(SORT, [arr([1, 2, 3]), 2])).toEqual(arr([1, 2, 3]));
		expect(call(SORT, [arr([1, 2, 3]), 4])).toBe(ERRORS.VALUE);
		expect(call(SORT, [arr([1, 2, 3]), 2, 1, true])).toBe(ERRORS.VALUE);
	});

	it('sorts references', () => {
		const ctx = new MockContext(
			block([
				[3, 'c'],
				[1, 'a'],
				[null, 'blank'],
				[2, 'b'],
			]),
		);
		expect(call(SORT, [ctx.ref('A1:B4')], ctx)).toEqual(
			arr([1, 'a'], [2, 'b'], [3, 'c'], [null, 'blank']),
		);
		expect(call(SORT, [ctx.ref('A1:B4'), ctx.ref('Z9')], ctx)).toBe(ERRORS.VALUE);
	});

	it('sorts 100k rows', () => {
		const n = 100_000;
		const rows = Array.from({ length: n }, (_, i) => [(i * 7919) % n, `r${i}`]);
		const started = Date.now();
		const out = call(SORT, [arr(...rows), 1, -1]) as ArrayValue;
		expect(Date.now() - started).toBeLessThan(2000);
		expect(out.values[0][0]).toBe(n - 1);
		expect(out.values[n - 1][0]).toBe(0);
	});
});

describe('SORTBY', () => {
	const names = arr(['a'], ['b'], ['c'], ['d']);

	it('sorts by a separate key column', () => {
		expect(call(SORTBY, [names, arr([3], [1], [4], [2])])).toEqual(arr(['b'], ['d'], ['a'], ['c']));
		expect(call(SORTBY, [names, arr([3], [1], [4], [2]), -1])).toEqual(
			arr(['c'], ['a'], ['d'], ['b']),
		);
	});

	it('sorts by several keys with their own orders', () => {
		const dept = arr(['x'], ['y'], ['x'], ['y']);
		const score = arr([1], [2], [3], [4]);
		expect(call(SORTBY, [names, dept, 1, score, -1])).toEqual(arr(['c'], ['a'], ['d'], ['b']));
		expect(call(SORTBY, [names, dept, -1, score])).toEqual(arr(['b'], ['d'], ['a'], ['c']));
		expect(call(SORTBY, [names, dept, MISSING, score, -1])).toEqual(
			arr(['c'], ['a'], ['d'], ['b']),
		);
	});

	it('sorts columns by a key row', () => {
		expect(call(SORTBY, [arr([1, 2, 3], [4, 5, 6]), arr([3, 1, 2])])).toEqual(
			arr([2, 3, 1], [5, 6, 4]),
		);
	});

	it('uses Excel order with blanks last and keeps ties stable', () => {
		const keys = arr([null], ['b'], [1], ['B']);
		expect(call(SORTBY, [names, keys])).toEqual(arr(['c'], ['b'], ['d'], ['a']));
		expect(call(SORTBY, [names, keys, -1])).toEqual(arr(['b'], ['d'], ['c'], ['a']));
	});

	it('validates by_array shapes and orders', () => {
		expect(call(SORTBY, [names, arr([1], [2], [3])])).toBe(ERRORS.VALUE);
		expect(call(SORTBY, [names, arr([1, 2], [3, 4], [5, 6], [7, 8])])).toBe(ERRORS.VALUE);
		expect(call(SORTBY, [arr([1, 2], [3, 4]), arr([1], [2]), 1, arr([1, 2])])).toBe(ERRORS.VALUE);
		expect(call(SORTBY, [names, arr([1], [2], [3], [4]), 0])).toBe(ERRORS.VALUE);
		expect(call(SORTBY, [names, arr([1], [2], [3], [4]), arr([1, -1])])).toBe(ERRORS.VALUE);
		expect(call(SORTBY, [names, arr([1], [2], [3], [4]), ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(SORTBY, [names, arr([1], [2], [3], [4]), '-1'])).toEqual(
			arr(['d'], ['c'], ['b'], ['a']),
		);
	});

	it('sorts references by another reference', () => {
		const ctx = new MockContext(
			block([
				['x', 30],
				['y', 10],
				['z', 20],
			]),
		);
		expect(call(SORTBY, [ctx.ref('A1:A3'), ctx.ref('B1:B3')], ctx)).toEqual(
			arr(['y'], ['z'], ['x']),
		);
	});
});

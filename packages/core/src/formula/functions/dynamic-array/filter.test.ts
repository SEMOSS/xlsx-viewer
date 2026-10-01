import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { MultiRefValue } from '../../types';
import { block } from '../lookup/test-utils/helpers';
import { FILTER } from './filter';

const DATA = arr(['Ann', 'East', 10], ['Bob', 'West', 20], ['Cy', 'East', 30], ['Di', 'North', 40]);
const MISSING = { type: 'missing' } as const;

describe('FILTER rows', () => {
	it('keeps rows whose include value is TRUE', () => {
		expect(call(FILTER, [DATA, arr([true], [false], [true], [false])])).toEqual(
			arr(['Ann', 'East', 10], ['Cy', 'East', 30]),
		);
		expect(call(FILTER, [DATA, arr([false], [false], [false], [true])])).toEqual(
			arr(['Di', 'North', 40]),
		);
	});

	it('reads numbers, blanks and TRUE / FALSE text as logicals', () => {
		expect(call(FILTER, [DATA, arr([1], [0], [-2.5], [null])])).toEqual(
			arr(['Ann', 'East', 10], ['Cy', 'East', 30]),
		);
		expect(call(FILTER, [DATA, arr(['TRUE'], ['false'], [false], [false])])).toEqual(
			arr(['Ann', 'East', 10]),
		);
	});

	it('rejects other text and returns errors found in include', () => {
		expect(call(FILTER, [DATA, arr([true], ['yes'], [true], [true])])).toBe(ERRORS.VALUE);
		expect(call(FILTER, [DATA, arr([true], [ERRORS.DIV0], [ERRORS.NA], [true])])).toBe(ERRORS.DIV0);
	});

	it('returns if_empty, or #CALC! without it, when nothing is kept', () => {
		const none = arr([false], [false], [false], [0]);
		expect(call(FILTER, [DATA, none])).toBe(ERRORS.CALC);
		expect(call(FILTER, [DATA, none, MISSING])).toBe(ERRORS.CALC);
		expect(call(FILTER, [DATA, none, null])).toBe(ERRORS.CALC);
		expect(call(FILTER, [DATA, none, 'none'])).toBe('none');
		expect(call(FILTER, [DATA, none, ''])).toBe('');
		expect(call(FILTER, [DATA, none, ERRORS.NA])).toBe(ERRORS.NA);
		const ctx = new MockContext({ Z1: 'empty' });
		expect(call(FILTER, [DATA, none, ctx.ref('Z1')], ctx)).toEqual(ctx.ref('Z1'));
		expect(call(FILTER, [DATA, none, arr([0, 0])], ctx)).toEqual(arr([0, 0]));
	});

	it('requires include to match the array', () => {
		expect(call(FILTER, [DATA, arr([true], [true], [true])])).toBe(ERRORS.VALUE);
		expect(call(FILTER, [DATA, arr([true, true], [true, true], [true, true], [true, true])])).toBe(
			ERRORS.VALUE,
		);
		expect(call(FILTER, [DATA, true])).toBe(ERRORS.VALUE);
	});
});

describe('FILTER columns and references', () => {
	it('keeps columns with a horizontal include', () => {
		const a = arr([1, 2, 3], [4, 5, 6]);
		expect(call(FILTER, [a, arr([true, false, true])])).toEqual(arr([1, 3], [4, 6]));
		expect(call(FILTER, [a, arr([0, 1, 0])])).toEqual(arr([2], [5]));
		expect(call(FILTER, [arr([1, 2, 3]), arr([true, false, true])])).toEqual(arr([1, 3]));
	});

	it('treats a single value as a 1x1 array', () => {
		expect(call(FILTER, [5, true])).toEqual(arr([5]));
		expect(call(FILTER, [5, false])).toBe(ERRORS.CALC);
		expect(call(FILTER, [arr([1, 2]), true])).toEqual(arr([1, 2]));
	});

	it('filters references and keeps blank cells as blank', () => {
		const ctx = new MockContext(
			block([
				['a', 1],
				['b', null],
				['c', 3],
			]),
		);
		expect(call(FILTER, [ctx.ref('A1:B3'), arr([false], [true], [true])], ctx)).toEqual(
			arr(['b', null], ['c', 3]),
		);
		expect(call(FILTER, [ctx.ref('A1:A3'), ctx.ref('B1:B3')], ctx)).toEqual(arr(['a'], ['c']));
	});

	it('propagates errors in place of the arrays', () => {
		const ctx = new MockContext({ A1: 1 });
		const union: MultiRefValue = { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('C3')] };
		expect(call(FILTER, [union, true], ctx)).toBe(ERRORS.VALUE);
		expect(call(FILTER, [arr([1]), union], ctx)).toBe(ERRORS.VALUE);
		expect(call(FILTER, [arr([ERRORS.NA], [2]), arr([true], [false])])).toEqual(arr([ERRORS.NA]));
	});

	it('filters 100k rows quickly', () => {
		const rows = Array.from({ length: 100_000 }, (_, i) => [i, i % 3 === 0 ? 'keep' : 'skip']);
		const include = arr(...rows.map((r) => [r[1] === 'keep']));
		const started = Date.now();
		const out = call(FILTER, [arr(...rows), include]);
		expect(Date.now() - started).toBeLessThan(1000);
		expect(out).toMatchObject({ kind: 'array', rows: 33_334, cols: 2 });
	});
});

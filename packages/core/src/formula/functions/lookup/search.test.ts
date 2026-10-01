import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg, Scalar } from '../../types';
import {
	arrayArg,
	isOmitted,
	isOmittedArray,
	optionalBoolean,
	optionalInteger,
	scalarArg,
} from './args';
import { compareSameType, exactMatcher, orderAgainst, regexMatcher } from './compare';
import type { Vec } from './grid';
import { gridColumnResult, gridRowResult, isFullHeight, sameExtent, toGrid } from './grid';
import { approximateMatch, binarySearchFirst, binarySearchLast, linearSearch } from './search';
import { readXModes, xsearch } from './xsearch';

function vec(values: Scalar[], length = values.length): Vec & { reads: number } {
	const v = {
		reads: 0,
		length,
		dataLength: values.length,
		get(i: number): Scalar {
			v.reads++;
			return i < values.length ? values[i] : null;
		},
	};
	return v;
}

const MISSING = { type: 'missing' } as const;

describe('argument helpers', () => {
	const args: FunctionArg[] = [1, null, MISSING, 'x', arr([null]), arr([1, 2])];

	it('recognises omitted arguments', () => {
		expect(isOmitted(args, 0)).toBe(false);
		expect(isOmitted(args, 1)).toBe(true);
		expect(isOmitted(args, 2)).toBe(true);
		expect(isOmitted(args, 9)).toBe(true);
		expect(isOmittedArray(args, 4)).toBe(true);
		expect(isOmittedArray(args, 5)).toBe(false);
		expect(isOmittedArray([arr([{ type: 'missing' } as unknown as Scalar])], 0)).toBe(true);
	});

	it('reads scalars, numbers and logicals with defaults', () => {
		expect(scalarArg(args, 2)).toBeNull();
		expect(scalarArg(args, 3)).toBe('x');
		expect(optionalInteger(args, 1, 7)).toBe(7);
		expect(optionalInteger([2.9], 0, 7)).toBe(2);
		expect(optionalInteger([-2.9], 0, 7)).toBe(-2);
		expect(optionalInteger(['x'], 0, 7)).toBe(ERRORS.VALUE);
		expect(optionalBoolean(args, 2, true)).toBe(true);
		expect(optionalBoolean([0], 0, true)).toBe(false);
		expect(optionalBoolean(['maybe'], 0, true)).toBe(ERRORS.VALUE);
	});

	it('wraps array parameters', () => {
		expect(arrayArg([5], 0)).toEqual(arr([5]));
		expect(arrayArg([ERRORS.CALC], 0)).toBe(ERRORS.CALC);
		expect(arrayArg([MISSING], 0)).toEqual(arr([null]));
		expect(arrayArg([], 0)).toEqual(arr([null]));
	});
});

describe('matchers', () => {
	it('matches exactly without type coercion', () => {
		const m = exactMatcher(5);
		expect([5, '5', true, null].map(m)).toEqual([true, false, false, false]);
		expect([true, 1].map(exactMatcher(true))).toEqual([true, false]);
		expect(exactMatcher(ERRORS.NA)(ERRORS.NA)).toBe(false);
	});

	it('compares text case-insensitively but accent-sensitively', () => {
		const m = exactMatcher('Straße');
		expect(['STRASSE', 'straße', 'STRAßE', 'Strasse'].map(m)).toEqual([false, true, true, false]);
		expect(exactMatcher('café')('cafe')).toBe(false);
		expect(exactMatcher('ÉCOLE')('école')).toBe(true);
	});

	it('applies wildcards only when asked', () => {
		expect(exactMatcher('a*', { wildcards: true })('abc')).toBe(true);
		expect(exactMatcher('a*')('abc')).toBe(false);
		expect(exactMatcher('a~*', { wildcards: true })('a*')).toBe(true);
		expect(exactMatcher('a~b', { wildcards: true })('a~b')).toBe(true);
		expect(exactMatcher('~~', { wildcards: true })('~')).toBe(true);
	});

	it('handles blank and empty lookup values per policy', () => {
		expect([null, '', 0].map(exactMatcher(null))).toEqual([false, true, false]);
		expect([null, '', 0].map(exactMatcher('', { blankMatchesBlank: true }))).toEqual([
			true,
			true,
			false,
		]);
	});

	it('builds regular-expression matchers', () => {
		const m = regexMatcher('^\\d+$');
		expect(typeof m).toBe('function');
		if (typeof m === 'function') {
			expect(['123', 123, 'a1', null, true, ERRORS.NA].map(m)).toEqual([
				true,
				true,
				false,
				false,
				false,
				false,
			]);
		}
		expect(regexMatcher('[')).toBe(ERRORS.VALUE);
		expect(regexMatcher(ERRORS.NA)).toBe(ERRORS.NA);
	});

	it('orders values of the lookup type only', () => {
		const o = orderAgainst(5);
		expect([o.comparable(3), o.comparable('3'), o.comparable(null)]).toEqual([true, false, false]);
		expect([o.compare(3), o.compare(5), o.compare(7)]).toEqual([-1, 0, 1]);
		const t = orderAgainst('m');
		expect(Math.sign(t.compare('A'))).toBe(-1);
		expect(t.compare('M')).toBe(0);
		expect(orderAgainst(null).comparable('')).toBe(true);
		expect(orderAgainst(true).compare(false)).toBe(-1);
		expect(orderAgainst(ERRORS.NA).comparable(1)).toBe(false);
		expect(compareSameType(1, 2)).toBe(-1);
		expect(compareSameType(true, false)).toBe(1);
		expect(compareSameType(1, 'a')).toBe(0);
	});
});

describe('search primitives', () => {
	it('searches linearly in both directions', () => {
		const v = vec([3, 1, 3, 2]);
		expect(linearSearch(v, (x) => x === 3)).toBe(0);
		expect(linearSearch(v, (x) => x === 3, true)).toBe(2);
		expect(linearSearch(v, (x) => x === 9)).toBe(-1);
	});

	it('considers the blank tail past the data only for blank matchers', () => {
		const v = vec([1, 2], 10);
		expect(linearSearch(v, (x) => x === null)).toBe(2);
		expect(linearSearch(v, (x) => x === null, true)).toBe(9);
		expect(linearSearch(v, (x) => x === 5)).toBe(-1);
	});

	it('binary-searches sorted data for first / last positions', () => {
		const v = vec([1, 2, 2, 2, 3, 5]);
		const o = orderAgainst(2);
		expect(binarySearchLast(v, o, (s) => s <= 0)).toBe(3);
		expect(binarySearchFirst(v, o, (s) => s >= 0)).toBe(1);
		expect(binarySearchFirst(vec([1, 2]), orderAgainst(9), (s) => s >= 0)).toBe(-1);
		expect(binarySearchLast(vec([]), o, (s) => s <= 0)).toBe(-1);
	});

	it('steps over incomparable values while probing', () => {
		const v = vec(['h', 1, null, ERRORS.NA, 4, 'x', 6, true]);
		expect(approximateMatch(v, 5)).toBe(4);
		expect(approximateMatch(v, 100)).toBe(6);
		expect(approximateMatch(v, 0)).toBe(-1);
		expect(approximateMatch(v, 'zz')).toBe(5);
		expect(approximateMatch(v, true)).toBe(7);
		expect(approximateMatch(v, ERRORS.NA)).toBe(-1);
	});

	it('finds the smallest value >= x in descending data', () => {
		const v = vec([9, 7, 7, 4, 1]);
		expect(approximateMatch(v, 5, true)).toBe(2);
		expect(approximateMatch(v, 7, true)).toBe(2);
		expect(approximateMatch(v, 10, true)).toBe(-1);
		expect(approximateMatch(v, 0, true)).toBe(4);
	});

	it('reads O(log n) cells on 100k sorted values', () => {
		const values = Array.from({ length: 100_000 }, (_, i) => i * 2);
		const v = vec(values);
		expect(approximateMatch(v, 54_321)).toBe(27_160);
		expect(v.reads).toBeLessThanOrEqual(17);
	});
});

describe('xsearch', () => {
	it('validates modes', () => {
		expect(readXModes([], 0, false)).toEqual({ match: 0, search: 1 });
		expect(readXModes([-1, -2], 0, false)).toEqual({ match: -1, search: -2 });
		expect(readXModes([3, 2], 0, false)).toEqual({ match: 3, search: 2 });
		expect(readXModes([4], 0, false)).toBe(ERRORS.VALUE);
		expect(readXModes([0, 0], 0, false)).toBe(ERRORS.VALUE);
		expect(readXModes([ERRORS.NA], 0, false)).toBe(ERRORS.NA);
	});

	it('dispatches every match / search combination', () => {
		const v = vec([10, 30, 20, 30]);
		expect(xsearch(v, 30, { match: 0, search: 1 })).toBe(1);
		expect(xsearch(v, 30, { match: 0, search: -1 })).toBe(3);
		expect(xsearch(v, 25, { match: -1, search: 1 })).toBe(2);
		expect(xsearch(v, 25, { match: 1, search: 1 })).toBe(1);
		expect(xsearch(v, 25, { match: 1, search: -1 })).toBe(3);
		expect(xsearch(vec(['ab', 'cd']), 'c?', { match: 2, search: 1 })).toBe(1);
		expect(xsearch(vec(['ab', 'cd']), 'c?', { match: 2, search: 2 })).toBe(ERRORS.VALUE);
		expect(xsearch(vec(['ab', 'cd']), '[', { match: 3, search: 1 })).toBe(ERRORS.VALUE);
	});

	it('prefers an exact blank match beyond the data over approximations', () => {
		expect(xsearch(vec(['a', 'b'], 5), '', { match: -1, search: 1 })).toBe(2);
		expect(xsearch(vec(['a', 'b'], 5), '', { match: 1, search: -1 })).toBe(4);
	});
});

describe('grids', () => {
	const ctx = new MockContext({ A1: 1, B1: 2, A2: 3, B2: 4, D9: 'far' });

	it('views references lazily with used-range data bounds', () => {
		const g = toGrid(ctx.ref('A:B'), ctx);
		if ('type' in g) {
			throw new Error('expected a grid');
		}
		expect([g.rows, g.cols, g.dataRows, g.dataCols]).toEqual([1_048_576, 2, 9, 2]);
		expect(g.get(1, 1)).toBe(4);
		expect(g.column(0).dataLength).toBe(9);
		expect(g.row(20).dataLength).toBe(0);
		expect(isFullHeight(g)).toBe(true);
		const empty = new MockContext();
		const e = toGrid(empty.ref('A1:C3'), empty);
		expect('type' in e ? null : [e.dataRows, e.dataCols]).toEqual([0, 0]);
	});

	it('wraps arrays and scalars and rejects bad inputs', () => {
		const a = toGrid(arr([1, 2]), ctx);
		expect('type' in a ? null : [a.rows, a.cols, a.ref]).toEqual([1, 2, null]);
		const s = toGrid('x', ctx);
		expect('type' in s ? null : s.get(0, 0)).toBe('x');
		expect(toGrid(ERRORS.DIV0, ctx)).toBe(ERRORS.DIV0);
		expect(toGrid(MISSING, ctx)).toBe(ERRORS.VALUE);
		expect(toGrid(undefined, ctx)).toBe(ERRORS.VALUE);
		const lambda = {
			kind: 'lambda' as const,
			params: [],
			body: { type: 'number' as const, value: 1 },
			scope: new Map(),
		};
		expect(toGrid(lambda, ctx)).toBe(ERRORS.CALC);
	});

	it('slices rows and columns as references or arrays', () => {
		const g = toGrid(ctx.ref('A1:B2'), ctx);
		const a = toGrid(arr([1, 2], [3, 4]), ctx);
		if ('type' in g || 'type' in a) {
			throw new Error('expected grids');
		}
		expect(gridRowResult(g, 1)).toEqual({
			kind: 'ref',
			sheet: 0,
			range: { top: 1, left: 0, bottom: 1, right: 1 },
		});
		expect(gridColumnResult(g, 1)).toEqual({
			kind: 'ref',
			sheet: 0,
			range: { top: 0, left: 1, bottom: 1, right: 1 },
		});
		expect(gridRowResult(a, 0)).toEqual(arr([1, 2]));
		expect(gridColumnResult(a, 1)).toEqual(arr([2], [4]));
		expect(gridRowResult(a, 5)).toBeNull();
		expect(gridColumnResult(a, 5)).toBeNull();
	});

	it('accepts clamped arrays against whole-column references', () => {
		const full = toGrid(ctx.ref('A:A'), ctx);
		const part = toGrid(ctx.ref('A1:A9'), ctx);
		const a = toGrid(arr([1], [2]), ctx);
		if ('type' in full || 'type' in part || 'type' in a) {
			throw new Error('expected grids');
		}
		expect(sameExtent(full, full.rows, a, a.rows, isFullHeight)).toBe(true);
		expect(sameExtent(a, a.rows, full, full.rows, isFullHeight)).toBe(true);
		expect(sameExtent(part, part.rows, a, a.rows, isFullHeight)).toBe(false);
		expect(sameExtent(full, full.rows, part, part.rows, isFullHeight)).toBe(false);
	});
});

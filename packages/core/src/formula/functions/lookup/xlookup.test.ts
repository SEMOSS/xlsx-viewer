import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { FormulaValue } from '../../types';
import { block, column, CountingContext, plain, row } from './test-utils/helpers';
import { XLOOKUP } from './xlookup';

// A1:D6; D4 is blank and code B2 appears twice.
const CELLS = block([
	['Code', 'Name', 'Price', 'Qty'],
	['A1', 'Apple', 1.5, 10],
	['B2', 'Banana', 0.5, 20],
	['C3', 'Cherry', 3, null],
	['B2', 'Blueberry', 4, 5],
	['D4', 'Date', 2, 7],
]);
const ctx = new MockContext(CELLS);
const codes = ctx.ref('A2:A6');
const names = ctx.ref('B2:B6');
const prices = ctx.ref('C2:C6');
const MISSING = { type: 'missing' } as const;

function x(args: unknown[], c: MockContext = ctx): unknown {
	return plain(call(XLOOKUP, args as FormulaValue[], c), c);
}

describe('XLOOKUP exact match', () => {
	it('returns a reference to the matching cell', () => {
		expect(call(XLOOKUP, ['C3', codes, names], ctx)).toEqual({
			kind: 'ref',
			sheet: 0,
			range: { top: 3, left: 1, bottom: 3, right: 1 },
		});
		expect(x(['C3', codes, names])).toBe('Cherry');
		expect(x(['c3', codes, prices])).toBe(3);
	});

	it('returns the first match by default and the last with search_mode -1', () => {
		expect(x(['B2', codes, names])).toBe('Banana');
		expect(x(['B2', codes, names, MISSING, 0, 1])).toBe('Banana');
		expect(x(['B2', codes, names, MISSING, 0, -1])).toBe('Blueberry');
	});

	it('uses if_not_found only when nothing matches', () => {
		expect(x(['Z9', codes, names])).toBe(ERRORS.NA);
		expect(x(['Z9', codes, names, 'none'])).toBe('none');
		expect(x(['Z9', codes, names, 0])).toBe(0);
		expect(x(['Z9', codes, names, ctx.ref('B1')])).toBe('Name');
		expect(x(['Z9', codes, names, MISSING])).toBe(ERRORS.NA);
		expect(x(['Z9', codes, names, null])).toBe(ERRORS.NA);
		expect(x(['Z9', codes, names, ERRORS.DIV0])).toBe(ERRORS.DIV0);
		expect(x(['A1', codes, ctx.ref('B2:B5'), 'none'])).toBe(ERRORS.VALUE);
	});

	it('never coerces types and treats * literally', () => {
		expect(x(['1.5', prices, names])).toBe(ERRORS.NA);
		expect(x([1.5, prices, names])).toBe('Apple');
		expect(x(['B*', codes, names])).toBe(ERRORS.NA);
		const lit = new MockContext(
			block([
				['a*', 1],
				['ab', 2],
			]),
		);
		expect(x(['a*', lit.ref('A1:A2'), lit.ref('B1:B2')], lit)).toBe(1);
	});

	it('matches a blank lookup value to the first blank or empty cell', () => {
		expect(x([null, ctx.ref('D2:D6'), names])).toBe('Cherry');
		expect(x(['', ctx.ref('D2:D6'), names])).toBe('Cherry');
		const tail = new MockContext({ ...column('A', [1, 2, 3]), ...column('B', ['a', 'b', 'c']) });
		const result = call(XLOOKUP, [null, tail.ref('A:A'), tail.ref('B:B')], tail);
		expect(result).toMatchObject({ range: { top: 3, bottom: 3 } });
		expect(plain(result, tail)).toBeNull();
		const last = call(XLOOKUP, [null, tail.ref('A:A'), tail.ref('B:B'), MISSING, 0, -1], tail);
		expect(last).toMatchObject({ range: { top: 1_048_575 } });
	});

	it('propagates errors from the lookup value and arrays', () => {
		expect(x([ERRORS.NUM, codes, names])).toBe(ERRORS.NUM);
		expect(x(['A1', ERRORS.REF, names])).toBe(ERRORS.REF);
		expect(x(['A1', codes, ERRORS.NA])).toBe(ERRORS.NA);
	});
});

describe('XLOOKUP shapes', () => {
	it('returns a whole row from a 2-D return_array', () => {
		expect(x(['B2', codes, ctx.ref('B2:D6')])).toEqual([['Banana', 0.5, 20]]);
		expect(x(['C3', codes, ctx.ref('B2:D6')])).toEqual([['Cherry', 3, null]]);
		expect(call(XLOOKUP, ['D4', codes, arr([1, 2], [3, 4], [5, 6], [7, 8], [9, 10])], ctx)).toEqual(
			arr([9, 10]),
		);
	});

	it('searches horizontally and returns a whole column', () => {
		const c = new MockContext(
			block([
				['Q1', 'Q2', 'Q3'],
				[100, 200, 300],
				[1, 2, 3],
			]),
		);
		expect(x(['Q2', c.ref('A1:C1'), c.ref('A2:C2')], c)).toBe(200);
		expect(x(['q3', c.ref('A1:C1'), c.ref('A2:C3')], c)).toEqual([[300], [3]]);
		expect(call(XLOOKUP, ['Q1', arr(['Q1', 'Q2']), arr([1, 2], [3, 4])], c)).toEqual(arr([1], [3]));
		expect(x(['Q2', c.ref('A1:C1'), c.ref('A2:B3')], c)).toBe(ERRORS.VALUE);
	});

	it('pairs a single-cell lookup_array with a row or a column', () => {
		expect(call(XLOOKUP, [1, arr([1]), arr([5, 6, 7])], ctx)).toEqual(arr([5, 6, 7]));
		expect(call(XLOOKUP, [1, arr([1]), arr([5], [6])], ctx)).toEqual(arr([5], [6]));
		expect(call(XLOOKUP, [1, arr([1]), arr([5, 6], [7, 8])], ctx)).toBe(ERRORS.VALUE);
		expect(call(XLOOKUP, [1, 1, 9], ctx)).toBe(9);
	});

	it('returns scalars from in-memory arrays', () => {
		expect(call(XLOOKUP, [2, arr([1], [2], [3]), arr(['a'], ['b'], ['c'])], ctx)).toBe('b');
	});

	it('rejects 2-D lookup arrays and mismatched sizes', () => {
		expect(x(['A1', ctx.ref('A2:B6'), names])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, ctx.ref('B2:B7')])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, ctx.ref('B1:F1')])).toBe(ERRORS.VALUE);
	});

	it('lifts over an array of lookup values', () => {
		expect(x([arr(['C3'], ['D4'], ['ZZ']), codes, prices])).toEqual([[3], [2], [ERRORS.NA]]);
		expect(x([arr(['A1', 'B2']), codes, names, '-'])).toEqual([['Apple', 'Banana']]);
	});

	it('accepts arrays computed from whole columns against whole-column references', () => {
		// B:B*2 or (A:A="x")*(B:B="y") arrive clamped to the used range (6 rows here).
		const doubled = arr(['x'], [3], [1], [6], [8], [4]);
		expect(x(['D4', ctx.ref('A:A'), doubled])).toBe(4);
		expect(x(['nope', ctx.ref('A:A'), doubled])).toBe(ERRORS.NA);
		const flags = arr([0], [0], [0], [1], [0], [0]);
		expect(x([1, flags, ctx.ref('B:B')])).toBe('Cherry');
		expect(x([1, flags, ctx.ref('B1:B10')])).toBe(ERRORS.VALUE);
		expect(x([1, arr([0], [1]), ctx.ref('B1:B6')])).toBe(ERRORS.VALUE);
	});

	it('reads a blank past the end of a clamped return array as blank', () => {
		const c = new MockContext({ ...column('A', [1, 2, 3, 4]), C10: 'far' });
		expect(x([4, c.ref('A:A'), arr(['a'], ['b'])], c)).toBeNull();
	});
});

describe('XLOOKUP approximate modes', () => {
	it('finds the next smaller or larger value in unsorted data', () => {
		expect(x([2.5, prices, names, MISSING, -1])).toBe('Date');
		expect(x([2.5, prices, names, MISSING, 1])).toBe('Cherry');
		expect(x([3, prices, names, MISSING, -1])).toBe('Cherry');
		expect(x([3, prices, names, MISSING, 1])).toBe('Cherry');
		expect(x([0.1, prices, names, MISSING, -1])).toBe(ERRORS.NA);
		expect(x([0.1, prices, names, MISSING, 1])).toBe('Banana');
		expect(x([10, prices, names, MISSING, 1])).toBe(ERRORS.NA);
		expect(x([10, prices, names, 'too big', 1])).toBe('too big');
		expect(x([10, prices, names, MISSING, -1])).toBe('Blueberry');
	});

	it('breaks ties by search direction', () => {
		const keys = arr([1], [3], [2], [3]);
		const vals = arr(['a'], ['b'], ['c'], ['d']);
		expect(call(XLOOKUP, [4, keys, vals, MISSING, -1], ctx)).toBe('b');
		expect(call(XLOOKUP, [4, keys, vals, MISSING, -1, -1], ctx)).toBe('d');
		expect(call(XLOOKUP, [2.5, keys, vals, MISSING, 1, -1], ctx)).toBe('d');
		expect(call(XLOOKUP, [3, keys, vals, MISSING, 1, -1], ctx)).toBe('d');
	});

	it('only considers values of the lookup type', () => {
		const keys = arr([5], ['x'], [true], [1]);
		const vals = arr(['five'], ['ex'], ['yes'], ['one']);
		expect(call(XLOOKUP, [3, keys, vals, MISSING, 1], ctx)).toBe('five');
		expect(call(XLOOKUP, [9, keys, vals, MISSING, 1], ctx)).toBe(ERRORS.NA);
		expect(call(XLOOKUP, ['m', keys, vals, MISSING, -1], ctx)).toBe(ERRORS.NA);
		expect(call(XLOOKUP, ['z', keys, vals, MISSING, -1], ctx)).toBe('ex');
		expect(call(XLOOKUP, [false, keys, vals, MISSING, 1], ctx)).toBe('yes');
	});

	it('compares text alphabetically', () => {
		const keys = arr(['delta'], ['Alpha'], ['charlie']);
		expect(call(XLOOKUP, ['bravo', keys, keys, MISSING, 1], ctx)).toBe('charlie');
		expect(call(XLOOKUP, ['bravo', keys, keys, MISSING, -1], ctx)).toBe('Alpha');
	});
});

describe('XLOOKUP wildcard and regex modes', () => {
	it('matches wildcards with match_mode 2', () => {
		expect(x(['B*', names, prices, MISSING, 2])).toBe(0.5);
		expect(x(['B*', names, prices, MISSING, 2, -1])).toBe(4);
		expect(x(['?pple', names, prices, MISSING, 2])).toBe(1.5);
		expect(x(['*rr*', names, prices, MISSING, 2])).toBe(3);
		expect(x(['x*', names, prices, 'none', 2])).toBe('none');
		const lit = new MockContext(
			block([
				['a*', 1],
				['ab', 2],
			]),
		);
		expect(x(['a~*', lit.ref('A1:A2'), lit.ref('B1:B2'), MISSING, 2], lit)).toBe(1);
		expect(x(['a?', lit.ref('A1:A2'), lit.ref('B1:B2'), MISSING, 2, -1], lit)).toBe(2);
	});

	it('matches regular expressions with match_mode 3', () => {
		expect(x(['^Ch', names, prices, MISSING, 3])).toBe(3);
		expect(x(['^ch', names, prices, MISSING, 3])).toBe(ERRORS.NA);
		expect(x(['err', names, prices, MISSING, 3])).toBe(3);
		expect(x(['err', names, prices, MISSING, 3, -1])).toBe(4);
		expect(x(['^[A-C]\\d$', codes, names, MISSING, 3, -1])).toBe('Blueberry');
		expect(x(['^0\\.5$', prices, names, MISSING, 3])).toBe('Banana');
		expect(x(['(', names, prices, MISSING, 3])).toBe(ERRORS.VALUE);
	});

	it('rejects wildcard or regex matching with a binary search', () => {
		expect(x(['B*', names, prices, MISSING, 2, 2])).toBe(ERRORS.VALUE);
		expect(x(['B', names, prices, MISSING, 3, -2])).toBe(ERRORS.VALUE);
	});
});

describe('XLOOKUP binary search', () => {
	const asc = arr([10], [20], [30], [40], [50]);
	const desc = arr([50], [40], [30], [20], [10]);
	const out = arr(['a'], ['b'], ['c'], ['d'], ['e']);

	it('searches ascending data with search_mode 2', () => {
		expect(call(XLOOKUP, [30, asc, out, MISSING, 0, 2], ctx)).toBe('c');
		expect(call(XLOOKUP, [35, asc, out, MISSING, 0, 2], ctx)).toBe(ERRORS.NA);
		expect(call(XLOOKUP, [35, asc, out, MISSING, -1, 2], ctx)).toBe('c');
		expect(call(XLOOKUP, [35, asc, out, MISSING, 1, 2], ctx)).toBe('d');
		expect(call(XLOOKUP, [5, asc, out, 'low', -1, 2], ctx)).toBe('low');
		expect(call(XLOOKUP, [55, asc, out, MISSING, 1, 2], ctx)).toBe(ERRORS.NA);
		expect(call(XLOOKUP, [55, asc, out, MISSING, -1, 2], ctx)).toBe('e');
		expect(call(XLOOKUP, [5, asc, out, MISSING, 1, 2], ctx)).toBe('a');
	});

	it('searches descending data with search_mode -2', () => {
		expect(call(XLOOKUP, [30, desc, out, MISSING, 0, -2], ctx)).toBe('c');
		expect(call(XLOOKUP, [35, desc, out, MISSING, 0, -2], ctx)).toBe(ERRORS.NA);
		expect(call(XLOOKUP, [35, desc, out, MISSING, -1, -2], ctx)).toBe('c');
		expect(call(XLOOKUP, [35, desc, out, MISSING, 1, -2], ctx)).toBe('b');
		expect(call(XLOOKUP, [60, desc, out, MISSING, -1, -2], ctx)).toBe('a');
		expect(call(XLOOKUP, [5, desc, out, MISSING, -1, -2], ctx)).toBe(ERRORS.NA);
	});

	it('finds the first of duplicates for an exact binary match', () => {
		const dup = arr([1], [2], [2], [2], [3]);
		expect(call(XLOOKUP, [2, dup, out, MISSING, 0, 2], ctx)).toBe('b');
		expect(call(XLOOKUP, [2, dup, out, MISSING, -1, 2], ctx)).toBe('d');
		expect(call(XLOOKUP, [2, dup, out, MISSING, 1, 2], ctx)).toBe('b');
	});

	it('searches text and skips other types', () => {
		const words = arr(['apple'], [7], ['Mango'], [null], ['pear']);
		expect(call(XLOOKUP, ['mango', words, out, MISSING, 0, 2], ctx)).toBe('c');
		expect(call(XLOOKUP, ['orange', words, out, MISSING, -1, 2], ctx)).toBe('c');
		expect(call(XLOOKUP, ['orange', words, out, MISSING, 1, 2], ctx)).toBe('e');
	});

	it('reads only a logarithmic number of cells on 100k rows', () => {
		const cells: Record<string, number | string> = {};
		for (let r = 1; r <= 100_000; r++) {
			cells[`A${r}`] = r;
			cells[`B${r}`] = `row${r}`;
		}
		const big = new CountingContext(cells);
		const keys = big.ref('A:A');
		const vals = big.ref('B:B');
		expect(plain(call(XLOOKUP, [77_777, keys, vals, MISSING, 0, 2], big), big)).toBe('row77777');
		expect(big.reads).toBeLessThanOrEqual(22);
		big.reads = 0;
		expect(plain(call(XLOOKUP, [77_777.5, keys, vals, MISSING, 1, 2], big), big)).toBe('row77778');
		expect(big.reads).toBeLessThanOrEqual(22);
		big.reads = 0;
		expect(plain(call(XLOOKUP, [99_999, keys, vals], big), big)).toBe('row99999');
		expect(big.reads).toBe(99_999 + 1);
		big.reads = 0;
		expect(plain(call(XLOOKUP, [99_999, keys, vals, MISSING, 0, -1], big), big)).toBe('row99999');
		expect(big.reads).toBe(2 + 1);
	});
});

describe('XLOOKUP mode validation', () => {
	it('rejects invalid match and search modes', () => {
		expect(x(['A1', codes, names, MISSING, 4])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, names, MISSING, -2])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, names, MISSING, 0, 0])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, names, MISSING, 0, 3])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, names, MISSING, 'x'])).toBe(ERRORS.VALUE);
		expect(x(['A1', codes, names, MISSING, ERRORS.NA])).toBe(ERRORS.NA);
	});

	it('truncates modes and defaults empty ones', () => {
		expect(x(['B2', codes, names, MISSING, 0.9, -1.5])).toBe('Blueberry');
		expect(x(['B2', codes, names, MISSING, MISSING, MISSING])).toBe('Banana');
		expect(x(['B2', codes, names, 'none', null, null])).toBe('Banana');
	});

	it('handles a horizontal reference row', () => {
		const c = new MockContext({ ...row(1, [1, 2, 3]), ...row(2, ['a', 'b', 'c']) });
		expect(x([3, c.ref('A1:C1'), c.ref('A2:C2')], c)).toBe('c');
		expect(x([3, c.ref('1:1'), c.ref('2:2')], c)).toBe('c');
	});
});

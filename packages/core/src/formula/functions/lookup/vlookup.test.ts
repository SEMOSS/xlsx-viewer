import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { MultiRefValue, RefValue } from '../../types';
import { block, CountingContext, plain } from './test-utils/helpers';
import { HLOOKUP, VLOOKUP } from './vlookup';

// A1:C6 with a text header row; C5 is blank.
const TABLE = block([
	['ID', 'Name', 'Score'],
	[1, 'Alice', 90],
	[2, 'bob', 85],
	[3, 'Carol', 70],
	[5, 'Dave', null],
	[8, 'Eve', 60],
]);

function ctxWith(extra: Record<string, unknown> = {}): MockContext {
	return new MockContext({ ...TABLE, ...(extra as Record<string, never>) });
}

describe('VLOOKUP exact match', () => {
	const ctx = ctxWith();
	const t = ctx.ref('A1:C6');

	it('finds numbers and returns the requested column', () => {
		expect(call(VLOOKUP, [3, t, 2, false], ctx)).toBe('Carol');
		expect(call(VLOOKUP, [3, t, 3, false], ctx)).toBe(70);
		expect(call(VLOOKUP, [8, t, 1, false], ctx)).toBe(8);
	});

	it('matches text case-insensitively', () => {
		expect(call(VLOOKUP, ['ALICE', ctx.ref('B1:C6'), 2, false], ctx)).toBe(90);
		expect(call(VLOOKUP, ['Bob', ctx.ref('B1:C6'), 2, false], ctx)).toBe(85);
		expect(call(VLOOKUP, ['name', ctx.ref('B1:C6'), 2, false], ctx)).toBe('Score');
	});

	it('supports * ? and ~ wildcards for text', () => {
		const t2 = ctx.ref('B1:C6');
		expect(call(VLOOKUP, ['c*', t2, 2, false], ctx)).toBe(70);
		expect(call(VLOOKUP, ['?ob', t2, 2, false], ctx)).toBe(85);
		expect(call(VLOOKUP, ['*e', t2, 2, false], ctx)).toBe('Score'); // the header "Name" comes first
		expect(call(VLOOKUP, ['*e', ctx.ref('B2:C6'), 2, false], ctx)).toBe(90);
		expect(call(VLOOKUP, ['D??e', t2, 1, false], ctx)).toBe('Dave');
		expect(call(VLOOKUP, ['z*', t2, 2, false], ctx)).toBe(ERRORS.NA);
		const lit = new MockContext({ A1: 'a*b', B1: 1, A2: 'axxb', B2: 2, A3: 'a?', B3: 3 });
		expect(call(VLOOKUP, ['a~*b', lit.ref('A1:B3'), 2, false], lit)).toBe(1);
		expect(call(VLOOKUP, ['a*b', lit.ref('A1:B3'), 2, false], lit)).toBe(1);
		expect(call(VLOOKUP, ['a~?', lit.ref('A1:B3'), 2, false], lit)).toBe(3);
	});

	it('returns #N/A when nothing matches', () => {
		expect(call(VLOOKUP, [4, t, 2, false], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, ['Zed', ctx.ref('B1:C6'), 2, false], ctx)).toBe(ERRORS.NA);
	});

	it('never coerces between numbers, text and logicals', () => {
		expect(call(VLOOKUP, ['1', t, 2, false], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [true, t, 2, false], ctx)).toBe(ERRORS.NA);
		const c = new MockContext({ A1: '7', B1: 'text', A2: 7, B2: 'number', A3: true, B3: 'bool' });
		expect(call(VLOOKUP, [7, c.ref('A1:B3'), 2, false], c)).toBe('number');
		expect(call(VLOOKUP, ['7', c.ref('A1:B3'), 2, false], c)).toBe('text');
		expect(call(VLOOKUP, [true, c.ref('A1:B3'), 2, false], c)).toBe('bool');
		expect(call(VLOOKUP, [1, c.ref('A1:B3'), 2, false], c)).toBe(ERRORS.NA);
	});

	it('returns blank (null) for an empty result cell', () => {
		expect(call(VLOOKUP, [5, t, 3, false], ctx)).toBeNull();
	});

	it('returns the first of duplicate keys', () => {
		const c = new MockContext(
			block([
				[1, 'a'],
				[2, 'b'],
				[2, 'c'],
				[3, 'd'],
			]),
		);
		expect(call(VLOOKUP, [2, c.ref('A1:B4'), 2, false], c)).toBe('b');
	});

	it('skips error cells in the key column', () => {
		const c = new MockContext(
			block([
				[ERRORS.NA, 'x'],
				[5, 'five'],
			]),
		);
		expect(call(VLOOKUP, [5, c.ref('A1:B2'), 2, false], c)).toBe('five');
	});

	it('matches empty text but never truly blank cells', () => {
		const c = new MockContext(
			block([
				[null, 'blank'],
				['', 'empty'],
				[0, 'zero'],
			]),
		);
		const t3 = c.ref('A1:B3');
		expect(call(VLOOKUP, ['', t3, 2, false], c)).toBe('empty');
		expect(call(VLOOKUP, [null, t3, 2, false], c)).toBe('empty');
		expect(call(VLOOKUP, [0, t3, 2, false], c)).toBe('zero');
		const noEmpty = new MockContext(
			block([
				[null, 'blank'],
				[0, 'zero'],
			]),
		);
		expect(call(VLOOKUP, [null, noEmpty.ref('A1:B2'), 2, false], noEmpty)).toBe(ERRORS.NA);
	});

	it('treats an empty range_lookup argument as FALSE', () => {
		expect(call(VLOOKUP, [4, t, 2, null], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [4, t, 2, { type: 'missing' }], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [4, t, 2, 0], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [4, t, 2, 'FALSE'], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [4, t, 2, 'abc'], ctx)).toBe(ERRORS.VALUE);
	});

	it('works on array constants', () => {
		const table = arr(['a', 1], ['b', 2], ['c', 3]);
		expect(call(VLOOKUP, ['B', table, 2, false], ctx)).toBe(2);
		expect(call(VLOOKUP, ['c', table, 1, false], ctx)).toBe('c');
		expect(call(VLOOKUP, ['bb', table, 2, true], ctx)).toBe(2);
	});
});

describe('VLOOKUP approximate match', () => {
	const ctx = ctxWith();
	const t = ctx.ref('A1:C6');

	it('defaults to approximate and finds the largest key <= value', () => {
		expect(call(VLOOKUP, [4, t, 2], ctx)).toBe('Carol');
		expect(call(VLOOKUP, [4, t, 2, true], ctx)).toBe('Carol');
		expect(call(VLOOKUP, [3, t, 2, 1], ctx)).toBe('Carol');
		expect(call(VLOOKUP, [7.99, t, 2], ctx)).toBe('Dave');
		expect(call(VLOOKUP, [100, t, 2], ctx)).toBe('Eve');
	});

	it('is #N/A below the smallest key', () => {
		expect(call(VLOOKUP, [0, t, 2], ctx)).toBe(ERRORS.NA);
		expect(call(VLOOKUP, [-5, ctx.ref('A2:C6'), 2], ctx)).toBe(ERRORS.NA);
	});

	it('ignores keys of another type (the text header)', () => {
		expect(call(VLOOKUP, [1, t, 2], ctx)).toBe('Alice');
		expect(call(VLOOKUP, ['ID', t, 2], ctx)).toBe('Name');
		expect(call(VLOOKUP, ['zzz', t, 2], ctx)).toBe('Name');
		expect(call(VLOOKUP, ['A', t, 2], ctx)).toBe(ERRORS.NA);
	});

	it('returns the last of duplicate keys', () => {
		const c = new MockContext(
			block([
				[1, 'a'],
				[2, 'b'],
				[2, 'c'],
				[2, 'd'],
				[3, 'e'],
			]),
		);
		expect(call(VLOOKUP, [2, c.ref('A1:B5'), 2, true], c)).toBe('d');
		expect(call(VLOOKUP, [2.5, c.ref('A1:B5'), 2, true], c)).toBe('d');
	});

	it('searches sorted text case-insensitively', () => {
		const c = new MockContext(
			block([
				['apple', 1],
				['Banana', 2],
				['cherry', 3],
			]),
		);
		expect(call(VLOOKUP, ['b', c.ref('A1:B3'), 2], c)).toBe(1);
		expect(call(VLOOKUP, ['BANANA', c.ref('A1:B3'), 2], c)).toBe(2);
		expect(call(VLOOKUP, ['zebra', c.ref('A1:B3'), 2], c)).toBe(3);
		expect(call(VLOOKUP, ['aardvark', c.ref('A1:B3'), 2], c)).toBe(ERRORS.NA);
	});

	it('does not apply wildcards', () => {
		const c = new MockContext(
			block([
				['a', 1],
				['b*', 2],
				['c', 3],
			]),
		);
		expect(call(VLOOKUP, ['b*', c.ref('A1:B3'), 2, true], c)).toBe(2);
	});

	it('reproduces binary-search answers on unsorted keys', () => {
		// Probes the middle (3 <= 5), then 9 > 5: Excel answers 3, not the exact 5.
		const c = new MockContext(
			block([
				[1, 'one'],
				[7, 'seven'],
				[3, 'three'],
				[9, 'nine'],
				[5, 'five'],
			]),
		);
		expect(call(VLOOKUP, [5, c.ref('A1:B5'), 2, true], c)).toBe('three');
		expect(call(VLOOKUP, [5, c.ref('A1:B5'), 2, false], c)).toBe('five');
		// The first key is above the lookup value but the middle probe still succeeds.
		const d = new MockContext(
			block([
				[6, 'six'],
				[1, 'one'],
				[2, 'two'],
				[3, 'three'],
				[4, 'four'],
			]),
		);
		expect(call(VLOOKUP, [5, d.ref('A1:B5'), 2, true], d)).toBe('four');
		// A large key in the middle hides everything after it.
		const e = new MockContext(
			block([
				[1, 'one'],
				[2, 'two'],
				[99, 'big'],
				[4, 'four'],
				[5, 'five'],
			]),
		);
		expect(call(VLOOKUP, [5, e.ref('A1:B5'), 2, true], e)).toBe('two');
	});

	it('steps over blanks and errors between sorted keys', () => {
		const c = new MockContext(
			block([
				[1, 'a'],
				[null, 'gap'],
				[3, 'c'],
				[ERRORS.DIV0, 'err'],
				[5, 'e'],
				[null, 'tail'],
			]),
		);
		expect(call(VLOOKUP, [4, c.ref('A1:B6'), 2], c)).toBe('c');
		expect(call(VLOOKUP, [9, c.ref('A1:B6'), 2], c)).toBe('e');
		expect(call(VLOOKUP, [2, c.ref('A1:B6'), 2], c)).toBe('a');
	});

	it('looks up logicals among logicals', () => {
		const c = new MockContext(
			block([
				[false, 'f'],
				[true, 't'],
			]),
		);
		expect(call(VLOOKUP, [true, c.ref('A1:B2'), 2], c)).toBe('t');
		expect(call(VLOOKUP, [false, c.ref('A1:B2'), 2], c)).toBe('f');
	});
});

describe('VLOOKUP arguments', () => {
	const ctx = ctxWith();
	const t = ctx.ref('A1:C6');

	it('validates col_index_num', () => {
		expect(call(VLOOKUP, [3, t, 0, false], ctx)).toBe(ERRORS.VALUE);
		expect(call(VLOOKUP, [3, t, -1, false], ctx)).toBe(ERRORS.VALUE);
		expect(call(VLOOKUP, [3, t, 4, false], ctx)).toBe(ERRORS.REF);
		expect(call(VLOOKUP, [99, t, 4, false], ctx)).toBe(ERRORS.REF);
		expect(call(VLOOKUP, [3, t, 'x', false], ctx)).toBe(ERRORS.VALUE);
		expect(call(VLOOKUP, [3, t, null, false], ctx)).toBe(ERRORS.VALUE);
	});

	it('truncates and coerces col_index_num', () => {
		expect(call(VLOOKUP, [3, t, 2.9, false], ctx)).toBe('Carol');
		expect(call(VLOOKUP, [3, t, '3', false], ctx)).toBe(70);
		expect(call(VLOOKUP, [3, t, true, false], ctx)).toBe(3);
	});

	it('propagates errors from arguments', () => {
		expect(call(VLOOKUP, [ERRORS.DIV0, t, 2, false], ctx)).toBe(ERRORS.DIV0);
		expect(call(VLOOKUP, [3, ERRORS.REF, 2, false], ctx)).toBe(ERRORS.REF);
		expect(call(VLOOKUP, [3, t, ERRORS.NUM, false], ctx)).toBe(ERRORS.NUM);
		expect(call(VLOOKUP, [3, t, 2, ERRORS.NA], ctx)).toBe(ERRORS.NA);
	});

	it('rejects lookup text longer than 255 characters', () => {
		expect(call(VLOOKUP, ['x'.repeat(256), t, 2, false], ctx)).toBe(ERRORS.VALUE);
		expect(call(VLOOKUP, ['x'.repeat(255), t, 2, false], ctx)).toBe(ERRORS.NA);
	});

	it('rejects unions and 3-D references', () => {
		const union: MultiRefValue = { kind: 'refs', refs: [ctx.ref('A1:B2'), ctx.ref('A4:B5')] };
		expect(call(VLOOKUP, [1, union, 2, false], ctx)).toBe(ERRORS.VALUE);
		const single: MultiRefValue = { kind: 'refs', refs: [t] };
		expect(call(VLOOKUP, [3, single, 2, false], ctx)).toBe('Carol');
		const threeD: RefValue = { ...ctx.ref('A1:C6'), sheetEnd: 1 };
		expect(call(VLOOKUP, [3, threeD, 2, false], ctx)).toBe(ERRORS.REF);
	});

	it('lifts over array lookup values and column indexes', () => {
		expect(call(VLOOKUP, [arr([1], [3], [4]), t, 2, false], ctx)).toEqual(
			arr(['Alice'], ['Carol'], [ERRORS.NA]),
		);
		expect(call(VLOOKUP, [2, t, arr([2, 3]), false], ctx)).toEqual(arr(['bob', 85]));
		const keys = new MockContext({ ...TABLE, E1: 2, E2: 8 });
		expect(call(VLOOKUP, [keys.ref('E1:E2'), keys.ref('A1:C6'), 3, false], keys)).toEqual(
			arr([85], [60]),
		);
	});

	it('intersects a range lookup value implicitly in legacy mode', () => {
		const c = new MockContext({ ...TABLE, E1: 1, E2: 3, E3: 8 }).at('F2');
		expect(call(VLOOKUP, [c.ref('E1:E3'), c.ref('A1:C6'), 2, false], c, 'legacy')).toBe('Carol');
	});

	it('reads whole-column tables lazily', () => {
		const cells: Record<string, number | string> = {};
		for (let r = 1; r <= 2000; r++) {
			cells[`A${r}`] = r * 2;
			cells[`B${r}`] = `v${r}`;
		}
		const c = new CountingContext(cells);
		const whole = c.ref('A:B');
		expect(call(VLOOKUP, [1001, whole, 2], c)).toBe('v500');
		expect(c.reads).toBeLessThan(40);
		c.reads = 0;
		expect(call(VLOOKUP, [20, whole, 2, false], c)).toBe('v10');
		expect(c.reads).toBe(11);
		c.reads = 0;
		expect(call(VLOOKUP, [3, whole, 2, false], c)).toBe(ERRORS.NA);
		expect(c.reads).toBe(2000);
	});
});

describe('HLOOKUP', () => {
	// A1:E3: keys across the first row.
	const ctx = new MockContext(
		block([
			['Key', 10, 20, 30, 40],
			['Name', 'ten', 'twenty', 'thirty', 'forty'],
			['Flag', true, false, null, true],
		]),
	);
	const t = ctx.ref('A1:E3');

	it('finds keys in the first row', () => {
		expect(call(HLOOKUP, [20, t, 2, false], ctx)).toBe('twenty');
		expect(call(HLOOKUP, [40, t, 3, false], ctx)).toBe(true);
		expect(call(HLOOKUP, ['key', t, 2, false], ctx)).toBe('Name');
		expect(call(HLOOKUP, [30, t, 3, false], ctx)).toBeNull();
	});

	it('does approximate matches along the row', () => {
		expect(call(HLOOKUP, [25, t, 2], ctx)).toBe('twenty');
		expect(call(HLOOKUP, [1000, t, 2, true], ctx)).toBe('forty');
		expect(call(HLOOKUP, [5, t, 2, true], ctx)).toBe(ERRORS.NA);
	});

	it('supports wildcards and validates row_index_num', () => {
		expect(call(HLOOKUP, ['K*', t, 2, false], ctx)).toBe('Name');
		expect(call(HLOOKUP, [20, t, 0, false], ctx)).toBe(ERRORS.VALUE);
		expect(call(HLOOKUP, [20, t, 4, false], ctx)).toBe(ERRORS.REF);
		expect(call(HLOOKUP, [21, t, 2, false], ctx)).toBe(ERRORS.NA);
	});

	it('works on array constants and lifts lookup values', () => {
		expect(call(HLOOKUP, ['b', arr(['a', 'b'], [1, 2]), 2, false], ctx)).toBe(2);
		expect(plain(call(HLOOKUP, [arr([10, 40]), t, 2, false], ctx), ctx)).toEqual([
			['ten', 'forty'],
		]);
	});
});

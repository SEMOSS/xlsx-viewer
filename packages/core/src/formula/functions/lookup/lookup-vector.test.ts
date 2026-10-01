import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { LOOKUP } from './lookup-vector';
import { block, column, row } from './test-utils/helpers';

describe('LOOKUP vector form', () => {
	const ctx = new MockContext({
		...column('A', [1, 3, 5, 7, 9]),
		...column('B', ['one', 'three', 'five', 'seven', 'nine']),
	});
	const keys = ctx.ref('A1:A5');
	const values = ctx.ref('B1:B5');

	it('returns the element at the matching position', () => {
		expect(call(LOOKUP, [5, keys, values], ctx)).toBe('five');
		expect(call(LOOKUP, [6, keys, values], ctx)).toBe('five');
		expect(call(LOOKUP, [1, keys, values], ctx)).toBe('one');
		expect(call(LOOKUP, [100, keys, values], ctx)).toBe('nine');
	});

	it('is #N/A below the smallest value', () => {
		expect(call(LOOKUP, [0, keys, values], ctx)).toBe(ERRORS.NA);
		expect(call(LOOKUP, ['a', keys, values], ctx)).toBe(ERRORS.NA);
	});

	it('extends a short result reference past its end', () => {
		expect(call(LOOKUP, [7, keys, ctx.ref('B1')], ctx)).toBe('seven');
		expect(call(LOOKUP, [9, keys, ctx.ref('B1:B2')], ctx)).toBe('nine');
	});

	it('accepts a result vector of the other orientation', () => {
		const c = new MockContext({ ...column('A', [1, 2, 3]), ...row(10, ['x', 'y', 'z']) });
		expect(call(LOOKUP, [2, c.ref('A1:A3'), c.ref('A10:C10')], c)).toBe('y');
		expect(call(LOOKUP, [3, c.ref('A1:A3'), arr(['p', 'q', 'r'])], c)).toBe('r');
	});

	it('extends a single result cell along the lookup direction', () => {
		const c = new MockContext({ ...row(1, [10, 20, 30]), ...row(2, ['a', 'b', 'c']) });
		expect(call(LOOKUP, [25, c.ref('A1:C1'), c.ref('A2')], c)).toBe('b');
		expect(call(LOOKUP, [30, c.ref('A1:C1'), c.ref('A2:C2')], c)).toBe('c');
	});

	it('is #N/A past the end of a result array and for a 2-D result', () => {
		expect(call(LOOKUP, [9, keys, arr(['a'], ['b'])], ctx)).toBe(ERRORS.NA);
		expect(call(LOOKUP, [3, keys, arr(['a', 'b'], ['c', 'd'])], ctx)).toBe(ERRORS.NA);
	});

	it('handles text lookups case-insensitively', () => {
		const words = arr(['apple'], ['Banana'], ['cherry']);
		expect(call(LOOKUP, ['BANANA', words, arr([1], [2], [3])])).toBe(2);
		expect(call(LOOKUP, ['blueberry', words, arr([1], [2], [3])])).toBe(2);
		expect(call(LOOKUP, ['zz', words, arr([1], [2], [3])])).toBe(3);
	});

	it('returns the last non-error with LOOKUP(2, 1/(...), ...)', () => {
		// 1/(A1:A7<>"") gives 1 for non-blank rows and #DIV/0! for blank ones.
		const d = ERRORS.DIV0;
		const flags = arr([1], [d], [1], [d], [d], [1], [d]);
		const vals = arr(['a'], [null], ['b'], [null], [null], ['c'], [null]);
		expect(call(LOOKUP, [2, flags, vals])).toBe('c');
		expect(call(LOOKUP, [2, arr([1], [d], [d], [d], [d], [d], [d]), vals])).toBe('a');
		expect(call(LOOKUP, [2, arr([d], [d], [d], [d], [d], [d], [1]), vals])).toBeNull();
		expect(call(LOOKUP, [2, arr([d], [d]), arr(['x'], ['y'])])).toBe(ERRORS.NA);
	});

	it('returns the last number or text with the big-value idioms', () => {
		const c = new MockContext(column('A', ['Total', 4, 'n/a', 8, null, 'note', 15, 'end']));
		expect(call(LOOKUP, [9.99e307, c.ref('A:A')], c)).toBe(15);
		expect(call(LOOKUP, ['z'.repeat(255), c.ref('A:A')], c)).toBe('end');
	});

	it('lifts over array lookup values and propagates errors', () => {
		expect(call(LOOKUP, [arr([2, 9]), keys, values], ctx)).toEqual(arr(['one', 'nine']));
		expect(call(LOOKUP, [ERRORS.NUM, keys, values], ctx)).toBe(ERRORS.NUM);
		expect(call(LOOKUP, [3, ERRORS.REF, values], ctx)).toBe(ERRORS.REF);
		expect(call(LOOKUP, [3, keys, ERRORS.NA], ctx)).toBe(ERRORS.NA);
	});
});

describe('LOOKUP array form', () => {
	it('searches the first column of a tall array and returns from the last column', () => {
		const table = arr([1, 'a', 'x'], [2, 'b', 'y'], [3, 'c', 'z']);
		expect(call(LOOKUP, [2, table])).toBe('y');
		expect(call(LOOKUP, [2.5, table])).toBe('y');
		expect(call(LOOKUP, [0, table])).toBe(ERRORS.NA);
	});

	it('searches the first row of a wide array and returns from the last row', () => {
		const table = arr([1, 2, 3, 4], ['a', 'b', 'c', 'd']);
		expect(call(LOOKUP, [3, table])).toBe('c');
		expect(call(LOOKUP, [10, table])).toBe('d');
	});

	it('treats a square array as tall', () => {
		expect(call(LOOKUP, [2, arr([1, 'p'], [2, 'q'])])).toBe('q');
	});

	it('works on a reference block', () => {
		const ctx = new MockContext(
			block([
				[10, 'ten'],
				[20, 'twenty'],
				[30, 'thirty'],
			]),
		);
		expect(call(LOOKUP, [25, ctx.ref('A1:B3')], ctx)).toBe('twenty');
		expect(call(LOOKUP, [10, ctx.ref('A1:A3')], ctx)).toBe(10);
	});

	it('looks up in a single row or column', () => {
		expect(call(LOOKUP, ['c', arr(['a', 'b', 'c'])])).toBe('c');
		expect(call(LOOKUP, [5, 5])).toBe(5);
	});
});

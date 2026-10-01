import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { ArrayValue } from '../../types';
import { column } from '../lookup/test-utils/helpers';
import { UNIQUE } from './unique';

describe('UNIQUE', () => {
	it('keeps the first occurrence of each value in order', () => {
		expect(call(UNIQUE, [arr([1], [2], [1], [3], [2])])).toEqual(arr([1], [2], [3]));
		expect(call(UNIQUE, [arr(['x'], ['y'], ['x'])])).toEqual(arr(['x'], ['y']));
	});

	it('compares text case-insensitively, keeping the first spelling', () => {
		expect(call(UNIQUE, [arr(['apple'], ['APPLE'], ['Apple'], ['pear'])])).toEqual(
			arr(['apple'], ['pear']),
		);
		expect(call(UNIQUE, [arr(['café'], ['cafe'])])).toEqual(arr(['café'], ['cafe']));
	});

	it('never merges values of different types', () => {
		const mixed = arr([1], ['1'], [true], [null], [''], [0], [false]);
		expect(call(UNIQUE, [mixed])).toEqual(mixed);
		expect(call(UNIQUE, [arr([0], [-0], [0.1 + 0.2], [0.3])])).toEqual(
			arr([0], [0.1 + 0.2], [0.3]),
		);
	});

	it('treats equal errors as duplicates', () => {
		expect(call(UNIQUE, [arr([ERRORS.NA], [ERRORS.NA], [ERRORS.DIV0], [1])])).toEqual(
			arr([ERRORS.NA], [ERRORS.DIV0], [1]),
		);
	});

	it('compares whole rows', () => {
		const rows = arr([1, 'a'], [1, 'A'], [1, 'b'], [2, 'a'], [1, 'b']);
		expect(call(UNIQUE, [rows])).toEqual(arr([1, 'a'], [1, 'b'], [2, 'a']));
		const tricky = arr(['1:sa', 'b'], ['1', 'sa:b'], ['1:sa', 'b']);
		expect(call(UNIQUE, [tricky])).toEqual(arr(['1:sa', 'b'], ['1', 'sa:b']));
	});

	it('compares columns with by_col', () => {
		const a = arr([1, 2, 1, 3], ['x', 'y', 'x', 'z']);
		expect(call(UNIQUE, [a, true])).toEqual(arr([1, 2, 3], ['x', 'y', 'z']));
		expect(call(UNIQUE, [a, true, true])).toEqual(arr([2, 3], ['y', 'z']));
		expect(call(UNIQUE, [arr([1, 1, 2]), true])).toEqual(arr([1, 2]));
		expect(call(UNIQUE, [arr([1, 1, 2])])).toEqual(arr([1, 1, 2]));
	});

	it('returns values that occur exactly once', () => {
		expect(call(UNIQUE, [arr([1], [2], [1], [3]), false, true])).toEqual(arr([2], [3]));
		expect(call(UNIQUE, [arr(['a'], ['A'], ['b']), false, true])).toEqual(arr(['b']));
		expect(call(UNIQUE, [arr([1], [1]), false, true])).toBe(ERRORS.CALC);
	});

	it('defaults omitted flags and validates them', () => {
		expect(call(UNIQUE, [arr([1], [1]), { type: 'missing' }, null])).toEqual(arr([1]));
		expect(call(UNIQUE, [arr([1], [1]), 'yes'])).toBe(ERRORS.VALUE);
		expect(call(UNIQUE, [arr([1], [1]), false, ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(UNIQUE, [arr([1], [1]), 'FALSE', 0])).toEqual(arr([1]));
	});

	it('reads references, keeping blanks as one distinct value', () => {
		const ctx = new MockContext(column('A', ['b', null, 'a', 'B', null, 'c']));
		expect(call(UNIQUE, [ctx.ref('A1:A6')], ctx)).toEqual(arr(['b'], [null], ['a'], ['c']));
		expect(call(UNIQUE, [7])).toEqual(arr([7]));
	});

	it('handles 100k rows quickly', () => {
		const values = arr(...Array.from({ length: 100_000 }, (_, i) => [`k${i % 1000}`]));
		const started = Date.now();
		const out = call(UNIQUE, [values]) as ArrayValue;
		expect(Date.now() - started).toBeLessThan(1000);
		expect(out.rows).toBe(1000);
		expect(out.values[999][0]).toBe('k999');
	});
});

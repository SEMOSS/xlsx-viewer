import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { filePrefixFor } from '../../prefixes';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { MultiRefValue } from '../../types';
import { block } from '../lookup/test-utils/helpers';
import { DYNAMIC_ARRAY_FUNCTIONS } from './index';
import { HSTACK, TOCOL, TOROW, VSTACK, WRAPCOLS, WRAPROWS } from './reshape';

const NA = ERRORS.NA;

describe('VSTACK / HSTACK', () => {
	it('stacks arrays vertically, padding narrower ones', () => {
		expect(call(VSTACK, [arr([1, 2]), arr([3, 4])])).toEqual(arr([1, 2], [3, 4]));
		expect(call(VSTACK, [arr([1, 2, 3]), arr([4]), 5])).toEqual(
			arr([1, 2, 3], [4, NA, NA], [5, NA, NA]),
		);
		expect(call(VSTACK, [arr([1])])).toEqual(arr([1]));
	});

	it('stacks arrays horizontally, padding shorter ones', () => {
		expect(call(HSTACK, [arr([1], [2]), arr([3], [4])])).toEqual(arr([1, 3], [2, 4]));
		expect(call(HSTACK, [arr([1], [2], [3]), arr([4, 5]), 'x'])).toEqual(
			arr([1, 4, 5, 'x'], [2, NA, NA, NA], [3, NA, NA, NA]),
		);
	});

	it('stacks references and propagates materialisation errors', () => {
		const ctx = new MockContext(
			block([
				[1, 2],
				[3, null],
			]),
		);
		expect(call(VSTACK, [ctx.ref('A1:B1'), ctx.ref('A2:B2')], ctx)).toEqual(arr([1, 2], [3, null]));
		expect(call(HSTACK, [ctx.ref('B1:B2'), ctx.ref('A1:A2')], ctx)).toEqual(arr([2, 1], [null, 3]));
		const union: MultiRefValue = { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('B2')] };
		expect(call(VSTACK, [arr([1]), union], ctx)).toBe(ERRORS.VALUE);
		expect(call(HSTACK, [ERRORS.DIV0, 1])).toEqual(arr([ERRORS.DIV0, 1]));
	});
});

describe('TOCOL / TOROW', () => {
	const a = arr([1, null, 3], [ERRORS.DIV0, 5, null]);

	it('scans by row or by column', () => {
		expect(call(TOCOL, [arr([1, 2], [3, 4])])).toEqual(arr([1], [2], [3], [4]));
		expect(call(TOCOL, [arr([1, 2], [3, 4]), 0, true])).toEqual(arr([1], [3], [2], [4]));
		expect(call(TOROW, [arr([1, 2], [3, 4])])).toEqual(arr([1, 2, 3, 4]));
		expect(call(TOROW, [arr([1, 2], [3, 4]), 0, true])).toEqual(arr([1, 3, 2, 4]));
	});

	it('skips blanks and / or errors', () => {
		expect(call(TOROW, [a])).toEqual(arr([1, null, 3, ERRORS.DIV0, 5, null]));
		expect(call(TOROW, [a, 1])).toEqual(arr([1, 3, ERRORS.DIV0, 5]));
		expect(call(TOROW, [a, 2])).toEqual(arr([1, null, 3, 5, null]));
		expect(call(TOROW, [a, 3])).toEqual(arr([1, 3, 5]));
		expect(call(TOCOL, [a, 3, true])).toEqual(arr([1], [5], [3]));
		expect(call(TOROW, [arr([''], [null]), 1])).toEqual(arr(['']));
	});

	it('validates ignore and reports empty results', () => {
		expect(call(TOCOL, [a, 4])).toBe(ERRORS.VALUE);
		expect(call(TOCOL, [a, -1])).toBe(ERRORS.VALUE);
		expect(call(TOCOL, [a, 'x'])).toBe(ERRORS.VALUE);
		expect(call(TOCOL, [arr([null, ERRORS.NA]), 3])).toBe(ERRORS.CALC);
		expect(call(TOCOL, [a, { type: 'missing' }, null])).toEqual(
			arr([1], [null], [3], [ERRORS.DIV0], [5], [null]),
		);
		expect(call(TOCOL, [7])).toEqual(arr([7]));
	});
});

describe('WRAPROWS / WRAPCOLS', () => {
	const seven = arr([1, 2, 3, 4, 5, 6, 7]);

	it('wraps a vector into rows, padding the last one', () => {
		expect(call(WRAPROWS, [seven, 3])).toEqual(arr([1, 2, 3], [4, 5, 6], [7, NA, NA]));
		expect(call(WRAPROWS, [seven, 3, 0])).toEqual(arr([1, 2, 3], [4, 5, 6], [7, 0, 0]));
		expect(call(WRAPROWS, [seven, 7])).toEqual(seven);
		expect(call(WRAPROWS, [arr([1], [2], [3]), 2])).toEqual(arr([1, 2], [3, NA]));
		expect(call(WRAPROWS, [arr([1, 2, 3]), 5])).toEqual(arr([1, 2, 3, NA, NA]));
	});

	it('wraps a vector into columns', () => {
		expect(call(WRAPCOLS, [seven, 3])).toEqual(arr([1, 4, 7], [2, 5, NA], [3, 6, NA]));
		expect(call(WRAPCOLS, [seven, 4, ''])).toEqual(arr([1, 5], [2, 6], [3, 7], [4, '']));
	});

	it('allows errors as padding and validates arguments', () => {
		expect(call(WRAPROWS, [arr([1, 2, 3]), 2, ERRORS.DIV0])).toEqual(arr([1, 2], [3, ERRORS.DIV0]));
		expect(call(WRAPROWS, [arr([1, 2], [3, 4]), 2])).toBe(ERRORS.VALUE);
		expect(call(WRAPROWS, [seven, 0])).toBe(ERRORS.NUM);
		expect(call(WRAPCOLS, [seven, -1])).toBe(ERRORS.NUM);
		expect(call(WRAPROWS, [seven, 2.9])).toEqual(arr([1, 2], [3, 4], [5, 6], [7, NA]));
		expect(call(WRAPROWS, [seven, ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(WRAPCOLS, [seven, 'x'])).toBe(ERRORS.VALUE);
	});
});

describe('DYNAMIC_ARRAY_FUNCTIONS', () => {
	it('registers every function with its file prefix and array flag', () => {
		expect(DYNAMIC_ARRAY_FUNCTIONS.map((d) => d.name).sort()).toEqual([
			'CHOOSECOLS',
			'CHOOSEROWS',
			'DROP',
			'EXPAND',
			'FILTER',
			'GROUPBY',
			'HSTACK',
			'PIVOTBY',
			'RANDARRAY',
			'SEQUENCE',
			'SORT',
			'SORTBY',
			'TAKE',
			'TOCOL',
			'TOROW',
			'UNIQUE',
			'VSTACK',
			'WRAPCOLS',
			'WRAPROWS',
		]);
		for (const def of DYNAMIC_ARRAY_FUNCTIONS) {
			expect(def.filePrefix).toBe(filePrefixFor(def.name));
			expect(def.returnsArray).toBe(true);
			expect(Boolean(def.volatile)).toBe(def.name === 'RANDARRAY');
		}
	});
});

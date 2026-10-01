import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { filePrefixFor } from '../../prefixes';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { LOOKUP_FUNCTIONS } from './index';
import { GETPIVOTDATA, HYPERLINK, TRANSPOSE } from './misc';
import { block } from './test-utils/helpers';

describe('TRANSPOSE', () => {
	it('swaps rows and columns of arrays', () => {
		expect(call(TRANSPOSE, [arr([1, 2, 3], [4, 5, 6])])).toEqual(arr([1, 4], [2, 5], [3, 6]));
		expect(call(TRANSPOSE, [arr([1, 2, 3])])).toEqual(arr([1], [2], [3]));
		expect(call(TRANSPOSE, [arr([1], [2])])).toEqual(arr([1, 2]));
	});

	it('transposes references, keeping blanks', () => {
		const ctx = new MockContext(
			block([
				['a', null],
				[true, 2],
			]),
		);
		expect(call(TRANSPOSE, [ctx.ref('A1:B2')], ctx)).toEqual(arr(['a', true], [null, 2]));
	});

	it('wraps scalars and propagates materialisation errors', () => {
		expect(call(TRANSPOSE, [7])).toEqual(arr([7]));
		expect(call(TRANSPOSE, [ERRORS.NA])).toEqual(arr([ERRORS.NA]));
		const ctx = new MockContext({ A1: 1 });
		const union = { kind: 'refs' as const, refs: [ctx.ref('A1'), ctx.ref('B2')] };
		expect(call(TRANSPOSE, [union], ctx)).toBe(ERRORS.VALUE);
	});
});

describe('HYPERLINK', () => {
	it('shows the friendly name, or the link when there is none', () => {
		expect(call(HYPERLINK, ['https://example.com', 'Example'])).toBe('Example');
		expect(call(HYPERLINK, ['https://example.com'])).toBe('https://example.com');
		expect(call(HYPERLINK, ['https://example.com', { type: 'missing' }])).toBe(
			'https://example.com',
		);
		expect(call(HYPERLINK, ['#Sheet2!A1', 42])).toBe(42);
		expect(call(HYPERLINK, ['https://example.com', null])).toBeNull();
	});

	it('rejects links longer than 255 characters and propagates errors', () => {
		expect(call(HYPERLINK, ['x'.repeat(256)])).toBe(ERRORS.VALUE);
		expect(call(HYPERLINK, [ERRORS.REF, 'x'])).toBe(ERRORS.REF);
		expect(call(HYPERLINK, ['x', ERRORS.NA])).toBe(ERRORS.NA);
	});

	it('lifts over arrays', () => {
		expect(call(HYPERLINK, [arr(['a', 'b']), arr(['A', 'B'])])).toEqual(arr(['A', 'B']));
	});
});

describe('GETPIVOTDATA', () => {
	it('is #REF! because pivot tables are not evaluated', () => {
		const ctx = new MockContext({ A1: 1 });
		expect(call(GETPIVOTDATA, ['Sales', ctx.ref('A1')], ctx)).toBe(ERRORS.REF);
		expect(call(GETPIVOTDATA, ['Sales', ctx.ref('A1'), 'Region', 'East'], ctx)).toBe(ERRORS.REF);
	});
});

describe('LOOKUP_FUNCTIONS', () => {
	it('registers every lookup function once', () => {
		const names = LOOKUP_FUNCTIONS.map((d) => d.name).sort();
		expect(names).toEqual([
			'GETPIVOTDATA',
			'HLOOKUP',
			'HYPERLINK',
			'LOOKUP',
			'MATCH',
			'TRANSPOSE',
			'VLOOKUP',
			'XLOOKUP',
			'XMATCH',
		]);
	});

	it('declares file prefixes consistent with the prefix table', () => {
		for (const def of LOOKUP_FUNCTIONS) {
			expect(def.filePrefix ?? '').toBe(filePrefixFor(def.name));
		}
	});

	it('flags array results and keeps lookup tables as references', () => {
		const byName = new Map(LOOKUP_FUNCTIONS.map((d) => [d.name, d]));
		expect(byName.get('XLOOKUP')?.returnsArray).toBe(true);
		expect(byName.get('TRANSPOSE')?.returnsArray).toBe(true);
		for (const name of ['VLOOKUP', 'HLOOKUP', 'MATCH', 'LOOKUP', 'XLOOKUP', 'XMATCH']) {
			expect(byName.get(name)?.params[1]).toBe('any');
		}
	});
});

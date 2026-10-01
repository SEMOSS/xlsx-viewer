import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 'Sale Price', A2: 'Sweden', A3: 2019 });
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);

describe('LEFT / RIGHT / MID', () => {
	it('extracts from the left', () => {
		expect(f('LEFT', 'Sale Price', 4)).toBe('Sale');
		expect(f('LEFT', 'Sweden')).toBe('S');
		expect(f('LEFT', 'abc', 10)).toBe('abc');
		expect(f('LEFT', 'abc', 0)).toBe('');
		expect(f('LEFT', 'abc', null)).toBe('');
		expect(f('LEFT', 1234.5, 3)).toBe('123');
		expect(f('LEFT', true, 2)).toBe('TR');
		expect(f('LEFT', 'abc', 1.9)).toBe('a');
		expect(f('LEFTB', 'abc', 2)).toBe('ab');
	});

	it('extracts from the right', () => {
		expect(f('RIGHT', 'Sale Price', 5)).toBe('Price');
		expect(f('RIGHT', 'Stock Number')).toBe('r');
		expect(f('RIGHT', 'abc', 0)).toBe('');
		expect(f('RIGHT', 'abc', 5)).toBe('abc');
		expect(f('RIGHTB', 'abc', 2)).toBe('bc');
	});

	it('extracts from the middle', () => {
		expect(f('MID', 'Fluid Flow', 1, 5)).toBe('Fluid');
		expect(f('MID', 'Fluid Flow', 7, 20)).toBe('Flow');
		expect(f('MID', 'Fluid Flow', 20, 5)).toBe('');
		expect(f('MID', 'abc', 2, 0)).toBe('');
		expect(f('MIDB', 'abcdef', 2, 3)).toBe('bcd');
	});

	it('rejects negative counts and positions before the start', () => {
		expect(f('LEFT', 'abc', -1)).toBe(ERRORS.VALUE);
		expect(f('RIGHT', 'abc', -1)).toBe(ERRORS.VALUE);
		expect(f('MID', 'abc', 0, 1)).toBe(ERRORS.VALUE);
		expect(f('MID', 'abc', 1, -1)).toBe(ERRORS.VALUE);
		expect(f('MID', 'abc', 1)).toBe(ERRORS.VALUE);
		expect(f('LEFT', ERRORS.NA, 1)).toBe(ERRORS.NA);
		expect(f('LEFT', 'abc', 'x')).toBe(ERRORS.VALUE);
	});

	it('lifts over ranges and arrays', () => {
		expect(f('LEFT', ctx.ref('A1:A3'), 2)).toEqual(arr(['Sa'], ['Sw'], ['20']));
		expect(f('MID', 'abcdef', arr([1, 3]), 2)).toEqual(arr(['ab', 'cd']));
	});
});

describe('REPLACE', () => {
	it('replaces by position', () => {
		expect(f('REPLACE', 'abcdefghijk', 6, 5, '*')).toBe('abcde*k');
		expect(f('REPLACE', '2009', 3, 2, '10')).toBe('2010');
		expect(f('REPLACE', '123456', 1, 3, '@')).toBe('@456');
		expect(f('REPLACE', 'abc', 5, 1, 'X')).toBe('abcX');
		expect(f('REPLACE', 'abc', 2, 0, 'X')).toBe('aXbc');
		expect(f('REPLACEB', 'abc', 1, 1, 'Z')).toBe('Zbc');
	});

	it('validates positions', () => {
		expect(f('REPLACE', 'abc', 0, 1, 'X')).toBe(ERRORS.VALUE);
		expect(f('REPLACE', 'abc', 1, -1, 'X')).toBe(ERRORS.VALUE);
	});
});

describe('SUBSTITUTE', () => {
	it('replaces every or the n-th occurrence', () => {
		expect(f('SUBSTITUTE', 'Sales Data', 'Sales', 'Cost')).toBe('Cost Data');
		expect(f('SUBSTITUTE', 'Quarter 1, 2008', '1', '2', 1)).toBe('Quarter 2, 2008');
		expect(f('SUBSTITUTE', 'Quarter 1, 2011', '1', '2', 3)).toBe('Quarter 1, 2012');
		expect(f('SUBSTITUTE', 'aaa', 'aa', 'b')).toBe('ba');
		expect(f('SUBSTITUTE', 'aaaa', 'aa', 'b', 2)).toBe('aab');
		expect(f('SUBSTITUTE', 1231, 1, 9)).toBe('9239');
	});

	it('leaves the text unchanged when nothing matches', () => {
		expect(f('SUBSTITUTE', 'abc', '', 'x')).toBe('abc');
		expect(f('SUBSTITUTE', 'abc', 'z', 'x')).toBe('abc');
		expect(f('SUBSTITUTE', 'abc', 'b', 'x', 5)).toBe('abc');
		expect(f('SUBSTITUTE', 'ABC', 'b', 'x')).toBe('ABC');
	});

	it('validates the instance number and result length', () => {
		expect(f('SUBSTITUTE', 'abc', 'b', 'x', 0)).toBe(ERRORS.VALUE);
		expect(f('SUBSTITUTE', 'abc', 'b', 'x', null)).toBe(ERRORS.VALUE);
		expect(f('SUBSTITUTE', 'a'.repeat(200), 'a', 'b'.repeat(200))).toBe(ERRORS.VALUE);
	});
});

describe('FIND / SEARCH', () => {
	it('finds case-sensitively without wildcards', () => {
		expect(f('FIND', 'M', 'Miriam McGovern')).toBe(1);
		expect(f('FIND', 'm', 'Miriam McGovern')).toBe(6);
		expect(f('FIND', 'M', 'Miriam McGovern', 3)).toBe(8);
		expect(f('FIND', '*', 'a*b')).toBe(2);
		expect(f('FIND', 1, ctx.ref('A3'))).toBe(3);
		expect(f('FINDB', 'c', 'abc')).toBe(3);
	});

	it('searches case-insensitively with wildcards', () => {
		expect(f('SEARCH', 'e', 'Statements', 6)).toBe(7);
		expect(f('SEARCH', 'margin', 'Profit Margin')).toBe(8);
		expect(f('SEARCH', 'p*t', 'Profit')).toBe(1);
		expect(f('SEARCH', '?r', 'Profit')).toBe(1);
		expect(f('SEARCH', 'B', 'abc')).toBe(2);
		expect(f('SEARCH', '~*', 'a*b')).toBe(2);
		expect(f('SEARCH', '~?', 'what?')).toBe(5);
		expect(f('SEARCH', '~', 'a~b')).toBe(2);
		expect(f('SEARCH', 'c*', 'abcabc', 4)).toBe(6);
		expect(f('SEARCH', '(', 'f(x)')).toBe(2);
		expect(f('SEARCHB', 'C', 'abc')).toBe(3);
	});

	it('handles empty find text and start positions', () => {
		expect(f('FIND', '', 'abc')).toBe(1);
		expect(f('FIND', '', 'abc', 3)).toBe(3);
		expect(f('SEARCH', '', 'abc', 2)).toBe(2);
		expect(f('FIND', 'a', 'abc', 0)).toBe(ERRORS.VALUE);
		expect(f('FIND', 'a', 'abc', 4)).toBe(ERRORS.VALUE);
		expect(f('FIND', 'a', 'abc', null)).toBe(ERRORS.VALUE);
		expect(f('FIND', '', '')).toBe(ERRORS.VALUE);
	});

	it('returns #VALUE! when not found and propagates errors', () => {
		expect(f('FIND', 'z', 'abc')).toBe(ERRORS.VALUE);
		expect(f('SEARCH', 'x', 'abc')).toBe(ERRORS.VALUE);
		expect(f('FIND', 'A', 'abc')).toBe(ERRORS.VALUE);
		expect(f('SEARCH', ERRORS.REF, 'abc')).toBe(ERRORS.REF);
	});

	it('lifts over arrays', () => {
		expect(f('FIND', 'a', arr(['abc', 'bca', 'xyz']))).toEqual(arr([1, 3, ERRORS.VALUE]));
	});
});

import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 'alpha', A2: 'Beta', A3: 12.5, B1: 'kana', B2: ERRORS.NA });
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);

describe('LEN / LENB', () => {
	it('counts UTF-16 code units of the text form', () => {
		expect(f('LEN', 'Phoenix, AZ')).toBe(11);
		expect(f('LEN', '')).toBe(0);
		expect(f('LEN', '     One   ')).toBe(11);
		expect(f('LEN', 123.45)).toBe(6);
		expect(f('LEN', true)).toBe(4);
		expect(f('LEN', null)).toBe(0);
		expect(f('LEN', 1 / 3)).toBe(17);
		expect(f('LEN', '\u{1F600}')).toBe(2);
		expect(f('LENB', 'abc')).toBe(3);
	});

	it('propagates errors and lifts over ranges', () => {
		expect(f('LEN', ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('LEN', ctx.ref('A1:A3'))).toEqual(arr([5], [4], [4]));
	});
});

describe('LOWER / UPPER / PROPER', () => {
	it('changes case', () => {
		expect(f('LOWER', 'E. E. Cummings')).toBe('e. e. cummings');
		expect(f('LOWER', 'Apt. 2B')).toBe('apt. 2b');
		expect(f('UPPER', 'total')).toBe('TOTAL');
		expect(f('UPPER', 'Yield')).toBe('YIELD');
		expect(f('UPPER', true)).toBe('TRUE');
		expect(f('LOWER', 123)).toBe('123');
		expect(f('UPPER', 'straße')).toBe('STRAßE');
		expect(f('UPPER', 'école')).toBe('ÉCOLE');
	});

	it('capitalises words like Excel', () => {
		expect(f('PROPER', 'this is a TITLE')).toBe('This Is A Title');
		expect(f('PROPER', '2-way street')).toBe('2-Way Street');
		expect(f('PROPER', '76BudGet')).toBe('76Budget');
		expect(f('PROPER', "don't")).toBe("Don'T");
		expect(f('PROPER', 'éCOLE à paris')).toBe('École À Paris');
		expect(f('PROPER', 'cafés')).toBe('Cafés');
		expect(f('PROPER', ctx.ref('A1:A2'))).toEqual(arr(['Alpha'], ['Beta']));
	});
});

describe('TRIM / CLEAN', () => {
	it('trims and collapses spaces only', () => {
		expect(f('TRIM', ' First Quarter Earnings ')).toBe('First Quarter Earnings');
		expect(f('TRIM', 'a   b    c')).toBe('a b c');
		expect(f('TRIM', '\ta ')).toBe('\ta ');
		expect(f('TRIM', '   ')).toBe('');
		expect(f('TRIM', 42)).toBe('42');
	});

	it('removes the first 32 control characters', () => {
		expect(f('CLEAN', '\tMonthly report\n')).toBe('Monthly report');
		expect(f('CLEAN', 'a\u0000b\u001fc')).toBe('abc');
		expect(f('CLEAN', 'a\u007fb')).toBe('a\u007fb');
	});
});

describe('EXACT', () => {
	it('compares case-sensitively', () => {
		expect(f('EXACT', 'word', 'word')).toBe(true);
		expect(f('EXACT', 'Word', 'word')).toBe(false);
		expect(f('EXACT', 'w ord', 'word')).toBe(false);
		expect(f('EXACT', 1, '1')).toBe(true);
		expect(f('EXACT', true, 'TRUE')).toBe(true);
		expect(f('EXACT', null, '')).toBe(true);
		expect(f('EXACT', 'a', ERRORS.DIV0)).toBe(ERRORS.DIV0);
	});
});

describe('REPT', () => {
	it('repeats text', () => {
		expect(f('REPT', '*-', 3)).toBe('*-*-*-');
		expect(f('REPT', '-', 10)).toBe('----------');
		expect(f('REPT', 'ab', 0)).toBe('');
		expect(f('REPT', 'ab', 2.9)).toBe('abab');
		expect(f('REPT', 1, 3)).toBe('111');
	});

	it('rejects negative counts and oversized results', () => {
		expect(f('REPT', 'a', -1)).toBe(ERRORS.VALUE);
		expect(f('REPT', 'ab', 20000)).toBe(ERRORS.VALUE);
		expect(f('REPT', 'a', 'x')).toBe(ERRORS.VALUE);
	});
});

describe('T', () => {
	it('returns text only', () => {
		expect(f('T', 'Rainfall')).toBe('Rainfall');
		expect(f('T', 19)).toBe('');
		expect(f('T', true)).toBe('');
		expect(f('T', null)).toBe('');
		expect(f('T', ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('T', ctx.ref('A2:A3'))).toEqual(arr(['Beta'], ['']));
	});
});

describe('PHONETIC', () => {
	it('returns the text of the top-left cell', () => {
		expect(f('PHONETIC', ctx.ref('B1'))).toBe('kana');
		expect(f('PHONETIC', ctx.ref('B1:B2'))).toBe('kana');
		expect(f('PHONETIC', ctx.ref('A3'))).toBe('');
		expect(f('PHONETIC', ctx.ref('B2'))).toBe(ERRORS.NA);
		expect(f('PHONETIC', { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('B1')] })).toBe(ERRORS.NA);
	});
});

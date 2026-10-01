import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { toFullWidth, toHalfWidth } from './charsets';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);

describe('CHAR / CODE (Windows-1252)', () => {
	it('maps codes to characters', () => {
		expect(f('CHAR', 65)).toBe('A');
		expect(f('CHAR', 33)).toBe('!');
		expect(f('CHAR', 10)).toBe('\n');
		expect(f('CHAR', 65.9)).toBe('A');
		expect(f('CHAR', '66')).toBe('B');
		expect(f('CHAR', 128)).toBe('€');
		expect(f('CHAR', 151)).toBe('\u2014');
		expect(f('CHAR', 233)).toBe('é');
		expect(f('CHAR', arr([72, 105]))).toEqual(arr(['H', 'i']));
	});

	it('rejects codes outside 1-255', () => {
		expect(f('CHAR', 0)).toBe(ERRORS.VALUE);
		expect(f('CHAR', 256)).toBe(ERRORS.VALUE);
		expect(f('CHAR', 'x')).toBe(ERRORS.VALUE);
	});

	it('maps the first character back to its code', () => {
		expect(f('CODE', 'A')).toBe(65);
		expect(f('CODE', '!')).toBe(33);
		expect(f('CODE', 'Alphabet')).toBe(65);
		expect(f('CODE', '€')).toBe(128);
		expect(f('CODE', 'é')).toBe(233);
		expect(f('CODE', '一')).toBe(63);
		expect(f('CODE', 1)).toBe(49);
		expect(f('CODE', true)).toBe(84);
		expect(f('CODE', '')).toBe(ERRORS.VALUE);
		expect(f('CODE', null)).toBe(ERRORS.VALUE);
	});

	it('round-trips every byte', () => {
		for (let i = 1; i <= 255; i++) {
			expect(f('CODE', f('CHAR', i) as string)).toBe(i);
		}
	});
});

describe('UNICHAR / UNICODE', () => {
	it('converts code points', () => {
		expect(f('UNICHAR', 66)).toBe('B');
		expect(f('UNICHAR', 32)).toBe(' ');
		expect(f('UNICHAR', 0x1f600)).toBe('\u{1F600}');
		expect(f('UNICODE', ' ')).toBe(32);
		expect(f('UNICODE', 'B')).toBe(66);
		expect(f('UNICODE', '\u{1F600}x')).toBe(128512);
		expect(f('UNICODE', 5)).toBe(53);
	});

	it('validates code points', () => {
		expect(f('UNICHAR', 0)).toBe(ERRORS.VALUE);
		expect(f('UNICHAR', 0x110000)).toBe(ERRORS.VALUE);
		expect(f('UNICHAR', 0xd800)).toBe(ERRORS.NA);
		expect(f('UNICODE', '')).toBe(ERRORS.VALUE);
		expect(fn(TEXT_FUNCTIONS, 'UNICHAR').filePrefix).toBe('_xlfn.');
	});
});

describe('ASC / DBCS', () => {
	it('converts full-width characters to half-width', () => {
		expect(f('ASC', 'ＡＢＣ')).toBe('ABC');
		expect(f('ASC', '　１')).toBe(' 1');
		expect(f('ASC', 'カタカナ')).toBe('ｶﾀｶﾅ');
		expect(f('ASC', 'ガ')).toBe('ｶﾞ');
		expect(f('ASC', 'パ')).toBe('ﾊﾟ');
		expect(f('ASC', 'ヴ')).toBe('ｳﾞ');
		expect(f('ASC', 'abc')).toBe('abc');
	});

	it('converts half-width characters to full-width', () => {
		expect(f('DBCS', 'ABC')).toBe('ＡＢＣ');
		expect(f('DBCS', ' ')).toBe('　');
		expect(f('DBCS', 'ｶﾞ')).toBe('ガ');
		expect(f('DBCS', 'ﾊﾟ')).toBe('パ');
		expect(f('DBCS', 'ｳﾞ')).toBe('ヴ');
		expect(f('DBCS', 'ｱﾞ')).toBe('ア゛');
		expect(f('DBCS', 12)).toBe('１２');
	});

	it('round-trips', () => {
		const text = 'Excel 2024 ｶﾞｲｼﾞﾝ!';
		expect(toHalfWidth(toFullWidth(text))).toBe(text);
	});
});

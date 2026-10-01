import { describe, expect, it } from 'vitest';

import { BUILTIN_NUMBER_FORMATS } from '../styles/builtin-number-formats';
import { parseFormatCode, tokenizeNumberFormat } from './index';

const types = (code: string) => tokenizeNumberFormat(code).map((s) => s.map((t) => t.type));
const roundTrip = (code: string) =>
	tokenizeNumberFormat(code)
		.map((s) => s.map((t) => t.raw).join(''))
		.join(';');

describe('tokenizeNumberFormat', () => {
	it('splits sections at unquoted semicolons only', () => {
		expect(tokenizeNumberFormat('0;-0;"zero";@')).toHaveLength(4);
		expect(tokenizeNumberFormat('"a;b"0')).toHaveLength(1);
		expect(tokenizeNumberFormat('\\;0')).toHaveLength(1);
		expect(tokenizeNumberFormat('0;')).toHaveLength(2);
		expect(tokenizeNumberFormat(';;;')).toHaveLength(4);
	});

	it('lexes number codes', () => {
		expect(types('#,##0.00%')).toEqual([
			['digit', 'comma', 'digit', 'digit', 'digit', 'point', 'digit', 'digit', 'percent'],
		]);
		expect(types('0.00E+00')).toEqual([
			['digit', 'point', 'digit', 'digit', 'exponent', 'digit', 'digit'],
		]);
		expect(types('# ?/8')).toEqual([['digit', 'literal', 'digit', 'slash', 'numeral']]);
		expect(types('_(* @_)')).toEqual([['space', 'fill', 'text', 'space']]);
		expect(types('_(*- @_)')).toEqual([['space', 'fill', 'literal', 'text', 'space']]);
	});

	it('reads quoted text, escapes, spaces and fills', () => {
		expect(tokenizeNumberFormat('"$"\\-_)*x')[0]).toEqual([
			{ type: 'literal', text: '$', raw: '"$"' },
			{ type: 'literal', text: '-', raw: '\\-' },
			{ type: 'space', char: ')', raw: '_)' },
			{ type: 'fill', char: 'x', raw: '*x' },
		]);
	});

	it('classifies bracket tags', () => {
		const code = '[Red][Color10][>=100][=<5][$€-407][$-F800][$-x-systime][h][mm][DBNum1]';
		expect(tokenizeNumberFormat(code)[0]).toEqual([
			{ type: 'color', hex: '#FF0000', raw: '[Red]' },
			{ type: 'color', index: 10, raw: '[Color10]' },
			{ type: 'condition', op: '>=', value: 100, raw: '[>=100]' },
			{ type: 'condition', op: '<=', value: 5, raw: '[=<5]' },
			{ type: 'locale', symbol: '€', lcid: 0x407, raw: '[$€-407]' },
			{ type: 'locale', symbol: '', lcid: 0xf800, system: 'date', raw: '[$-F800]' },
			{ type: 'locale', symbol: '', system: 'time', raw: '[$-x-systime]' },
			{ type: 'elapsed', unit: 'h', length: 1, raw: '[h]' },
			{ type: 'elapsed', unit: 'm', length: 2, raw: '[mm]' },
			{ type: 'ignored', raw: '[DBNum1]' },
		]);
		expect(tokenizeNumberFormat('[Color57]')[0][0].type).toBe('ignored');
		expect(tokenizeNumberFormat('[<-1.5]')[0][0]).toMatchObject({ op: '<', value: -1.5 });
		expect(tokenizeNumberFormat('[$USD]')[0][0]).toMatchObject({ symbol: 'USD' });
	});

	it('lexes date codes, AM/PM and General case-insensitively', () => {
		const [s] = tokenizeNumberFormat('YYYY-mm-DD hh:MM:ss');
		expect(s.filter((t) => t.type === 'date').map((t) => t.raw)).toEqual([
			'YYYY',
			'mm',
			'DD',
			'hh',
			'MM',
			'ss',
		]);
		expect(tokenizeNumberFormat('am/pm')[0][0]).toMatchObject({ type: 'ampm', am: 'am', pm: 'pm' });
		expect(tokenizeNumberFormat('AM/PM')[0][0]).toMatchObject({ type: 'ampm', am: 'AM', pm: 'PM' });
		expect(tokenizeNumberFormat('A/P')[0][0]).toMatchObject({ type: 'ampm', am: 'A', pm: 'P' });
		expect(tokenizeNumberFormat('GENERAL')[0][0].type).toBe('general');
		expect(tokenizeNumberFormat('e')[0][0]).toMatchObject({ type: 'date', part: 'e' });
		expect(tokenizeNumberFormat('0e-0')[0][1]).toMatchObject({ type: 'exponent', upper: false });
		expect(tokenizeNumberFormat('B2yyyy')[0][0].type).toBe('ignored');
		expect(tokenizeNumberFormat('t0.00')[0][0].type).toBe('ignored');
	});

	it('round-trips every builtin code through token raws', () => {
		for (const code of Object.values(BUILTIN_NUMBER_FORMATS)) {
			expect(roundTrip(code)).toBe(code);
		}
	});

	it('round-trips malformed and unusual codes', () => {
		const codes = ['0"abc', '[Red 0', '0\\', '_', '*', '\\😀0', '0"😀"', '[$-F800]dddd;@'];
		for (const code of codes) {
			expect(roundTrip(code)).toBe(code);
		}
		expect(tokenizeNumberFormat('\\😀')[0][0]).toMatchObject({ type: 'literal', text: '😀' });
	});
});

describe('parseFormatCode', () => {
	it('caches parses per code', () => {
		expect(parseFormatCode('#,##0.00')).toBe(parseFormatCode('#,##0.00'));
	});

	it('recognises General', () => {
		expect(parseFormatCode('General').isGeneral).toBe(true);
		expect(parseFormatCode('general').isGeneral).toBe(true);
		expect(parseFormatCode('').isGeneral).toBe(true);
		expect(parseFormatCode('[Red]General').isGeneral).toBe(false);
	});

	it('separates numeric and text sections', () => {
		const four = parseFormatCode('0;-0;0;@');
		expect([four.numeric.length, four.text !== null]).toEqual([3, true]);
		const trailing = parseFormatCode('0.00;@');
		expect([trailing.numeric.length, trailing.text !== null]).toEqual([1, true]);
		const textOnly = parseFormatCode('@');
		expect([textOnly.numeric.length, textOnly.text !== null]).toEqual([0, true]);
		expect(parseFormatCode('0;-0').text).toBeNull();
	});

	it.each([
		['General', 'general'],
		['0.00', 'number'],
		['"abc"', 'number'],
		['0.00E+00', 'scientific'],
		['# ?/?', 'fraction'],
		['m/d/yyyy', 'date'],
		['[h]:mm', 'date'],
		['[Red]', 'general'],
		['0 "e"', 'number'],
	])('section kind of %s is %s', (code, kind) => {
		expect(parseFormatCode(code).sections[0].kind).toBe(kind);
	});

	it('compiles layouts', () => {
		expect(parseFormatCode('#,##0.00').numeric[0]).toMatchObject({
			intSlots: 4,
			decSlots: 2,
			grouping: true,
			scale: 0,
		});
		expect(parseFormatCode('0.0,,').numeric[0].scale).toBe(-6);
		expect(parseFormatCode('0%').numeric[0]).toMatchObject({ scale: 2, percent: true });
		expect(parseFormatCode('# ??/100').numeric[0]).toMatchObject({ denFixed: 100, numSlots: 2 });
		expect(parseFormatCode('[$€-407]0').numeric[0]).toMatchObject({
			currencySymbol: '€',
			lcid: 0x407,
		});
		expect(parseFormatCode('[Red]0').numeric[0].color).toBe('#FF0000');
		expect(parseFormatCode('[Color10]0').numeric[0].colorIndex).toBe(10);
		expect(parseFormatCode('h:mm:ss.00 AM/PM').numeric[0]).toMatchObject({
			hour12: true,
			subsecDigits: 2,
			hasTime: true,
			hasDate: false,
		});
		expect(parseFormatCode('[h]:mm').numeric[0].elapsedUnit).toBe('h');
	});

	it('derives section rules', () => {
		expect(parseFormatCode('0;0;0').rules.map((r) => r.dropSign)).toEqual([false, true, false]);
		const cond = parseFormatCode('[>=100]0;[<0]0;0');
		expect(cond.conditional).toBe(true);
		expect(cond.rules.map((r) => r.dropSign)).toEqual([false, true, false]);
		expect(parseFormatCode('[>100]0;0').rules[1]).toEqual({ dropSign: false });
		expect(parseFormatCode('0;[<-5]0;0').rules[0].condition).toEqual({ op: '>', value: 0 });
	});

	it('flags date and time-only formats', () => {
		expect(parseFormatCode('m/d/yyyy').isDate).toBe(true);
		expect(parseFormatCode('h:mm').isTimeOnly).toBe(true);
		expect(parseFormatCode('m/d/yyyy h:mm').isTimeOnly).toBe(false);
		expect(parseFormatCode('0.00').isDate).toBe(false);
	});
});

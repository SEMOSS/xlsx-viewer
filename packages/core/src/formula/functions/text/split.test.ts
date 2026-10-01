import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 'Red riding hood', A2: 'a,b', B1: ',', B2: ';' });
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);
const RIDING = "Little Red Riding Hood's red hood";

describe('TEXTBEFORE / TEXTAFTER', () => {
	it('extracts around the n-th delimiter', () => {
		expect(f('TEXTBEFORE', RIDING, 'Red')).toBe('Little ');
		expect(f('TEXTBEFORE', RIDING, 'red')).toBe("Little Red Riding Hood's ");
		expect(f('TEXTBEFORE', RIDING, 'red', 2)).toBe(ERRORS.NA);
		expect(f('TEXTBEFORE', RIDING, 'red', 2, 1)).toBe("Little Red Riding Hood's ");
		expect(f('TEXTBEFORE', RIDING, 'red', -1)).toBe("Little Red Riding Hood's ");
		expect(f('TEXTAFTER', RIDING, 'hood')).toBe('');
		expect(f('TEXTAFTER', RIDING, 'Hood')).toBe("'s red hood");
		expect(f('TEXTAFTER', RIDING, 'red', 2, 1)).toBe(' hood');
		expect(f('TEXTAFTER', RIDING, 'red', -2, 1)).toBe(" Riding Hood's red hood");
		expect(f('TEXTAFTER', RIDING, 'red', -2)).toBe(ERRORS.NA);
	});

	it('matches case-insensitively and against the end of the text', () => {
		expect(f('TEXTBEFORE', RIDING, 'RED', 1, 1)).toBe('Little ');
		expect(f('TEXTAFTER', RIDING, 'HOOD', 1, 1)).toBe("'s red hood");
		expect(f('TEXTBEFORE', 'a-b', '-', 2, 0, 1)).toBe('a-b');
		expect(f('TEXTAFTER', 'a-b', '-', 2, 0, 1)).toBe('');
		expect(f('TEXTBEFORE', 'abc', '-', 1, 0, 1)).toBe('abc');
		expect(f('TEXTAFTER', 'a-b', '-', -2, 0, 1)).toBe('a-b');
		expect(f('TEXTBEFORE', 'a-b', '-', -2, 0, 1)).toBe('');
	});

	it('accepts several delimiters', () => {
		expect(f('TEXTBEFORE', 'a;b,c', arr([',', ';']))).toBe('a');
		expect(f('TEXTAFTER', 'a;b,c', arr([',', ';']), 2)).toBe('c');
		expect(f('TEXTAFTER', 'a;b,c', ctx.ref('B1:B2'), -1)).toBe('c');
	});

	it('handles empty delimiters and empty text', () => {
		expect(f('TEXTBEFORE', 'abc', '')).toBe('');
		expect(f('TEXTAFTER', 'abc', '')).toBe('abc');
		expect(f('TEXTBEFORE', 'abc', '', -1)).toBe('abc');
		expect(f('TEXTAFTER', 'abc', '', -1)).toBe('');
		expect(f('TEXTBEFORE', 'abc', '', 2)).toBe('a');
		expect(f('TEXTBEFORE', '', ',')).toBe('');
	});

	it('reports missing delimiters and invalid instances', () => {
		expect(f('TEXTBEFORE', 'abc', 'x')).toBe(ERRORS.NA);
		expect(f('TEXTAFTER', 'abc', 'x', 1, 0, 0, 'none')).toBe('none');
		expect(f('TEXTBEFORE', 'a-b', '-', 2)).toBe(ERRORS.NA);
		expect(f('TEXTBEFORE', 'abc', 'b', 0)).toBe(ERRORS.VALUE);
		expect(f('TEXTBEFORE', 'abc', 'b', 4)).toBe(ERRORS.VALUE);
		expect(f('TEXTBEFORE', 'abc', 'b', 1, 2)).toBe(ERRORS.VALUE);
		expect(f('TEXTBEFORE', ERRORS.NA, 'b')).toBe(ERRORS.NA);
	});

	it('treats explicitly empty optional arguments as defaults', () => {
		expect(f('TEXTBEFORE', 'A,b', 'B', null, 1)).toBe('A,');
		expect(f('TEXTAFTER', 'abc', 'x', null, null, null, null)).toBe(ERRORS.NA);
	});

	it('lifts over the text argument', () => {
		expect(f('TEXTAFTER', ctx.ref('A1:A2'), arr([' ', ',']))).toEqual(arr(['riding hood'], ['b']));
		expect(fn(TEXT_FUNCTIONS, 'TEXTAFTER').filePrefix).toBe('_xlfn.');
	});
});

describe('TEXTSPLIT', () => {
	it('splits into columns and rows', () => {
		expect(f('TEXTSPLIT', 'Dakota Lennon Sanchez', ' ')).toEqual(
			arr(['Dakota', 'Lennon', 'Sanchez']),
		);
		expect(f('TEXTSPLIT', '1,2,3;4,5,6', ',', ';')).toEqual(arr(['1', '2', '3'], ['4', '5', '6']));
		expect(f('TEXTSPLIT', 'a;b', null, ';')).toEqual(arr(['a'], ['b']));
		expect(f('TEXTSPLIT', 'Do. Or do not. There is no try. -Anonymous', '.')).toEqual(
			arr(['Do', ' Or do not', ' There is no try', ' -Anonymous']),
		);
	});

	it('pads ragged rows', () => {
		expect(f('TEXTSPLIT', '1,2,3;4,5', ',', ';')).toEqual(
			arr(['1', '2', '3'], ['4', '5', ERRORS.NA]),
		);
		expect(f('TEXTSPLIT', '1,2,3;4,5', ',', ';', false, 0, '-')).toEqual(
			arr(['1', '2', '3'], ['4', '5', '-']),
		);
	});

	it('handles empty values, case and multiple delimiters', () => {
		expect(f('TEXTSPLIT', 'a,,b', ',')).toEqual(arr(['a', '', 'b']));
		expect(f('TEXTSPLIT', 'a,,b', ',', null, true)).toEqual(arr(['a', 'b']));
		expect(f('TEXTSPLIT', 'aXbxc', 'x', null, false, 1)).toEqual(arr(['a', 'b', 'c']));
		expect(f('TEXTSPLIT', 'aXbxc', 'x')).toEqual(arr(['aXb', 'c']));
		expect(f('TEXTSPLIT', 'a-b_c', arr(['-', '_']))).toEqual(arr(['a', 'b', 'c']));
		expect(f('TEXTSPLIT', 'a--b', arr(['-', '--']))).toEqual(arr(['a', 'b']));
		expect(f('TEXTSPLIT', '', ',')).toEqual(arr(['']));
	});

	it('reports invalid input', () => {
		expect(f('TEXTSPLIT', 'abc', '')).toBe(ERRORS.VALUE);
		expect(f('TEXTSPLIT', ',', ',', null, true)).toBe(ERRORS.CALC);
		expect(f('TEXTSPLIT', 'a', ',', null, false, 3)).toBe(ERRORS.VALUE);
		expect(f('TEXTSPLIT', ERRORS.REF, ',')).toBe(ERRORS.REF);
		expect(fn(TEXT_FUNCTIONS, 'TEXTSPLIT').returnsArray).toBe(true);
	});

	it('keeps only the first value per cell when lifted over a range', () => {
		expect(f('TEXTSPLIT', ctx.ref('A1:A2'), arr([' ', ',']))).toEqual(arr(['Red'], ['a']));
	});
});

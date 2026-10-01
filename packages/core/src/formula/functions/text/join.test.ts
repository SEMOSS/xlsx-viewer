import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';

const ctx = new MockContext({
	A1: 'a',
	A2: 'b',
	A3: 'c',
	B1: 'd',
	B3: 'f',
	C1: 1.5,
	C2: true,
	D1: ERRORS.DIV0,
	E1: 'Seattle',
	E2: 1234.01234,
	F1: true,
	F2: ERRORS.VALUE,
});
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);
const missing: AstNode = { type: 'missing' };

describe('CONCAT / CONCATENATE', () => {
	it('concatenates ranges row by row', () => {
		expect(f('CONCAT', ctx.ref('A1:B3'))).toBe('adbcf');
		expect(f('CONCAT', 'x', ctx.ref('C1:C2'), 3)).toBe('x1.5TRUE3');
		expect(f('CONCAT', arr(['p', 'q'], ['r', 's']))).toBe('pqrs');
		expect(f('CONCAT', 'a', missing, 'b')).toBe('ab');
		expect(fn(TEXT_FUNCTIONS, 'CONCAT').filePrefix).toBe('_xlfn.');
	});

	it('concatenates scalars and lifts over ranges', () => {
		expect(f('CONCATENATE', 'Stream population for ', 'brook trout', ' is ', 32, '/mile')).toBe(
			'Stream population for brook trout is 32/mile',
		);
		expect(f('CONCATENATE', true, null, 1.25)).toBe('TRUE1.25');
		expect(f('CONCATENATE', ctx.ref('A1:A2'), '!')).toEqual(arr(['a!'], ['b!']));
	});

	it('propagates errors and enforces the length limit', () => {
		expect(f('CONCAT', ctx.ref('C1:D1'))).toBe(ERRORS.DIV0);
		expect(f('CONCATENATE', 'a', ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('CONCAT', 'x'.repeat(20000), 'y'.repeat(20000))).toBe(ERRORS.VALUE);
	});
});

describe('TEXTJOIN', () => {
	it('joins with a delimiter', () => {
		expect(f('TEXTJOIN', ', ', true, 'US Dollar', 'Australian Dollar', 'Chinese Yuan')).toBe(
			'US Dollar, Australian Dollar, Chinese Yuan',
		);
		expect(f('TEXTJOIN', '-', true, ctx.ref('A1:B3'))).toBe('a-d-b-c-f');
		expect(f('TEXTJOIN', '-', false, ctx.ref('A1:B3'))).toBe('a-d-b--c-f');
		expect(f('TEXTJOIN', '', true, 'a', 'b')).toBe('ab');
		expect(f('TEXTJOIN', '', true, ...new Array<FunctionArg>(252).fill('x'))).toBe('x'.repeat(252));
		expect(f('TEXTJOIN', '', true, ...new Array<FunctionArg>(253).fill('x'))).toBe(ERRORS.VALUE);
		expect(f('TEXTJOIN', 1, true, 'a', 'b')).toBe('a1b');
	});

	it('uses several delimiters in turn', () => {
		expect(f('TEXTJOIN', arr(['-', '/']), true, 'a', 'b', 'c', 'd')).toBe('a-b/c-d');
		expect(f('TEXTJOIN', ctx.ref('A1:A2'), true, 1, 2, 3)).toBe('1a2b3');
	});

	it('handles empty values and errors', () => {
		expect(f('TEXTJOIN', ',', false, 'a', '', 'b')).toBe('a,,b');
		expect(f('TEXTJOIN', ',', true, 'a', '', 'b')).toBe('a,b');
		expect(f('TEXTJOIN', ',', null, 'a', missing, 'b')).toBe('a,,b');
		expect(f('TEXTJOIN', ',', true, ctx.ref('C1:D1'))).toBe(ERRORS.DIV0);
		expect(f('TEXTJOIN', ',', 'maybe', 'a')).toBe(ERRORS.VALUE);
		expect(f('TEXTJOIN', arr([ERRORS.NA]), true, 'a', 'b')).toBe(ERRORS.NA);
		expect(f('TEXTJOIN', ',', true, 'x'.repeat(20000), 'y'.repeat(20000))).toBe(ERRORS.VALUE);
	});
});

describe('VALUETOTEXT / ARRAYTOTEXT', () => {
	it('converts values in concise and strict formats', () => {
		expect(f('VALUETOTEXT', 'Hello')).toBe('Hello');
		expect(f('VALUETOTEXT', 'Hello', 1)).toBe('"Hello"');
		expect(f('VALUETOTEXT', 'say "hi"', 1)).toBe('"say ""hi"""');
		expect(f('VALUETOTEXT', 1234.01234, 1)).toBe('1234.01234');
		expect(f('VALUETOTEXT', true)).toBe('TRUE');
		expect(f('VALUETOTEXT', ERRORS.NA)).toBe('#N/A');
		expect(f('VALUETOTEXT', null)).toBe('');
		expect(f('VALUETOTEXT', 'x', 2)).toBe(ERRORS.VALUE);
		expect(f('VALUETOTEXT', 'x', ERRORS.REF)).toBe(ERRORS.REF);
		expect(f('VALUETOTEXT', ctx.ref('E1:E2'), 1)).toEqual(arr(['"Seattle"'], ['1234.01234']));
	});

	it('converts arrays in concise and strict formats', () => {
		expect(f('ARRAYTOTEXT', ctx.ref('E1:F2'))).toBe('Seattle, TRUE, 1234.01234, #VALUE!');
		expect(f('ARRAYTOTEXT', ctx.ref('E1:F2'), 1)).toBe('{"Seattle",TRUE;1234.01234,#VALUE!}');
		expect(f('ARRAYTOTEXT', arr([1, 2], [3, 4]), 1)).toBe('{1,2;3,4}');
		expect(f('ARRAYTOTEXT', 'x')).toBe('x');
		expect(f('ARRAYTOTEXT', arr([1]), 3)).toBe(ERRORS.VALUE);
		expect(fn(TEXT_FUNCTIONS, 'ARRAYTOTEXT').filePrefix).toBe('_xlfn.');
	});
});

import { describe, expect, it } from 'vitest';

import { ERRORS } from '../values/errors';
import { collectNumbers, compareScalars, numberToText, toBoolean, toNumber } from './coerce';
import { MockContext } from './test-utils/mock-context';
import { createCriteriaMatcher } from './criteria';
import { addFormulaPrefixes, stripFormulaPrefixes } from './prefixes';
import { translateFormula } from './translate';

describe('numberToText', () => {
	it.each([
		[0, '0'],
		[1, '1'],
		[-2.5, '-2.5'],
		[0.1 + 0.2, '0.3'],
		[123456789012345, '123456789012345'],
		[1e15, '1E+15'],
		[1234567890123456789, '1.23456789012346E+18'],
		[0.000123, '0.000123'],
		[1e-10, '1E-10'],
		[1 / 3, '0.333333333333333'],
	])('%d -> %s', (n, text) => {
		expect(numberToText(n)).toBe(text);
	});
});

describe('coercion', () => {
	it('coerces to numbers', () => {
		expect(toNumber(null)).toBe(0);
		expect(toNumber(true)).toBe(1);
		expect(toNumber(' 12 ')).toBe(12);
		expect(toNumber('50%')).toBe(0.5);
		expect(toNumber('abc')).toBe(ERRORS.VALUE);
		expect(toNumber('')).toBe(ERRORS.VALUE);
	});

	it('coerces to booleans', () => {
		expect(toBoolean('true')).toBe(true);
		expect(toBoolean(0)).toBe(false);
		expect(toBoolean('yes')).toBe(ERRORS.VALUE);
	});

	it('orders numbers < text < booleans, text case-insensitively', () => {
		expect(compareScalars(100, 'a')).toBeLessThan(0);
		expect(compareScalars('a', false)).toBeLessThan(0);
		expect(compareScalars('abc', 'ABC')).toBe(0);
		expect(compareScalars(null, 0)).toBe(0);
		expect(compareScalars(null, '')).toBe(0);
	});
});

describe('criteria', () => {
	const values = [1, 5, 10, 'apple', 'Banana', 'apricot', null, '', true, ERRORS.NA, '5'];
	const count = (c: Parameters<typeof createCriteriaMatcher>[0]) =>
		values.filter(createCriteriaMatcher(c)).length;

	it('matches numbers and numeric comparisons', () => {
		expect(count(5)).toBe(2); // 5 and "5"
		expect(count('>4')).toBe(2);
		expect(count('<=5')).toBe(2);
		expect(count('<>5')).toBe(values.length - 2);
	});

	it('matches text with wildcards, case-insensitively', () => {
		expect(count('ap*')).toBe(2);
		expect(count('banana')).toBe(1);
		expect(count('?pple')).toBe(1);
		expect(count('<>ap*')).toBe(values.length - 2);
	});

	it('matches blanks, non-blanks, booleans and errors', () => {
		expect(count('=')).toBe(1);
		expect(count('')).toBe(2);
		expect(count('<>')).toBe(values.length - 2);
		expect(count(true)).toBe(1);
		expect(count('#N/A')).toBe(1);
	});
});

describe('formula prefixes', () => {
	it('strips and restores future-function prefixes outside strings', () => {
		const file = '_xlfn.XLOOKUP(A1,B:B,C:C)&"_xlfn.KEEP"&_xlfn._xlws.FILTER(A1:A5,B1:B5>0)';
		const display = stripFormulaPrefixes(file);
		expect(display).toBe('XLOOKUP(A1,B:B,C:C)&"_xlfn.KEEP"&FILTER(A1:A5,B1:B5>0)');
		expect(addFormulaPrefixes(display)).toBe(file);
		expect(addFormulaPrefixes('SUM(A1:A3)')).toBe('SUM(A1:A3)');
	});
});

describe('translateFormula (interim)', () => {
	it('shifts relative references only', () => {
		expect(translateFormula('A1+$B$2+C$3+$D4', 1, 1)).toBe('B2+$B$2+D$3+$D5');
		expect(translateFormula('SUM(A1:A3)&"A1"', 2, 0)).toBe('SUM(A3:A5)&"A1"');
		expect(translateFormula("'Sheet A1'!B2", 1, 0)).toBe("'Sheet A1'!B3");
		expect(translateFormula('A1', -1, 0)).toBe('#REF!');
	});
});

describe('collectNumbers', () => {
	it('counts an omitted argument as 0, like Excel (AVERAGE(1,) is 0.5)', () => {
		expect(collectNumbers([1, { type: 'missing' }], new MockContext())).toEqual([1, 0]);
	});
});

describe('text comparison', () => {
	it('ignores case but not accents, like Excel', async () => {
		const { compareText } = await import('./coerce');
		expect(compareText('abc', 'ABC')).toBe(0);
		expect(compareText('a', 'á')).not.toBe(0);
		expect(compareText('a', 'b')).toBeLessThan(0);
	});
});

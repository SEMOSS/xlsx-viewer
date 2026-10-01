import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { TEXT_FUNCTIONS } from './index';
import { compilePattern, translatePattern } from './regex-compile';

const ctx = new MockContext({ A1: 'abc123', A2: 'xyz' });
const f = (name: string, ...args: FunctionArg[]) => call(fn(TEXT_FUNCTIONS, name), args, ctx);

describe('PCRE translation', () => {
	it('translates PCRE-only syntax', () => {
		expect(translatePattern('[[:digit:]]+')).toBe('[0-9]+');
		expect(translatePattern('(?P<year>\\d{4})-(?P=year)')).toBe('(?<year>\\d{4})-\\k<year>');
		expect(translatePattern('\\Aab\\z')).toBe('(?<![\\s\\S])ab(?![\\s\\S])');
		expect(translatePattern('[]a]')).toBe('[\\]a]');
		expect(translatePattern('[[:bogus:]]')).toBeNull();
		expect(translatePattern('[abc')).toBeNull();
	});

	it('compiles inline flags and rejects invalid patterns', () => {
		expect(compilePattern('(?i)abc', false, false)?.flags).toContain('i');
		expect(compilePattern('(?s)a.b', false, false)?.test('a\nb')).toBe(true);
		expect(compilePattern('(?x)a b', false, false)).toBeNull();
		expect(compilePattern('(', false, false)).toBeNull();
		expect(compilePattern('\\_', false, false)?.test('_')).toBe(true);
	});
});

describe('REGEXTEST', () => {
	it('tests patterns', () => {
		expect(f('REGEXTEST', 'alpha', 'a')).toBe(true);
		expect(f('REGEXTEST', 'alpha', '^l')).toBe(false);
		expect(f('REGEXTEST', '(555) 123-4567', '\\(\\d{3}\\) \\d{3}-\\d{4}')).toBe(true);
		expect(f('REGEXTEST', 'ALPHA', 'alpha')).toBe(false);
		expect(f('REGEXTEST', 'ALPHA', 'alpha', 1)).toBe(true);
		expect(f('REGEXTEST', 'ALPHA', '(?i)alpha')).toBe(true);
		expect(f('REGEXTEST', 12345, '^\\d+$')).toBe(true);
		expect(f('REGEXTEST', ctx.ref('A1:A2'), '\\d')).toEqual(arr([true], [false]));
	});

	it('rejects invalid patterns and options', () => {
		expect(f('REGEXTEST', 'a', '(')).toBe(ERRORS.VALUE);
		expect(f('REGEXTEST', 'a', 'a', 2)).toBe(ERRORS.VALUE);
		expect(f('REGEXTEST', ERRORS.NA, 'a')).toBe(ERRORS.NA);
		expect(fn(TEXT_FUNCTIONS, 'REGEXTEST').filePrefix).toBe('_xlfn.');
	});
});

describe('REGEXEXTRACT', () => {
	it('extracts the first match, all matches or capture groups', () => {
		expect(f('REGEXEXTRACT', 'DylanWilliams', '[A-Z][a-z]+')).toBe('Dylan');
		expect(f('REGEXEXTRACT', 'DylanWilliams', '[A-Z][a-z]+', 1)).toEqual(
			arr(['Dylan'], ['Williams']),
		);
		expect(f('REGEXEXTRACT', 'DylanWilliams', '([A-Z][a-z]+)([A-Z][a-z]+)', 2)).toEqual(
			arr(['Dylan', 'Williams']),
		);
		expect(f('REGEXEXTRACT', 'ab', '(a)(x)?b', 2)).toEqual(arr(['a', '']));
		expect(f('REGEXEXTRACT', 'abc', 'B', 0, 1)).toBe('b');
		expect(f('REGEXEXTRACT', 'abc', 'b', 2)).toEqual(arr(['b']));
	});

	it('returns #N/A without a match and validates options', () => {
		expect(f('REGEXEXTRACT', 'abc', '\\d')).toBe(ERRORS.NA);
		expect(f('REGEXEXTRACT', 'abc', '\\d', 1)).toBe(ERRORS.NA);
		expect(f('REGEXEXTRACT', 'abc', 'a', 3)).toBe(ERRORS.VALUE);
		expect(f('REGEXEXTRACT', 'abc', '[')).toBe(ERRORS.VALUE);
	});
});

describe('REGEXREPLACE', () => {
	it('replaces every match or a chosen occurrence', () => {
		expect(f('REGEXREPLACE', '555-123-4567', '[0-9]{3}-[0-9]{3}-([0-9]{4})', '***-***-$1')).toBe(
			'***-***-4567',
		);
		expect(f('REGEXREPLACE', 'a1b2c3', '\\d', '#')).toBe('a#b#c#');
		expect(f('REGEXREPLACE', 'a1b2c3', '\\d', '#', 2)).toBe('a1b#c3');
		expect(f('REGEXREPLACE', 'a1b2c3', '\\d', '#', -1)).toBe('a1b2c#');
		expect(f('REGEXREPLACE', 'a1b2c3', '\\d', '#', 5)).toBe('a1b2c3');
		expect(f('REGEXREPLACE', 'a1b2c3', '\\d', '#', -5)).toBe('a1b2c3');
		expect(f('REGEXREPLACE', 'ABC', 'b', 'x', 0, 1)).toBe('AxC');
	});

	it('expands group references', () => {
		expect(f('REGEXREPLACE', 'John Smith', '(\\w+) (\\w+)', '$2, $1')).toBe('Smith, John');
		expect(f('REGEXREPLACE', 'John Smith', '(?<first>\\w+) (\\w+)', '${2} ${first}')).toBe(
			'Smith John',
		);
		expect(f('REGEXREPLACE', 'cost', 'cost', '$$5 ($0)')).toBe('$5 (cost)');
		expect(f('REGEXREPLACE', 'abc', '', '-')).toBe('-a-b-c-');
	});

	it('validates arguments', () => {
		expect(f('REGEXREPLACE', 'abc', '(', 'x')).toBe(ERRORS.VALUE);
		expect(f('REGEXREPLACE', 'abc', 'b', ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('REGEXREPLACE', 'abc', 'b', 'x', 'y')).toBe(ERRORS.VALUE);
	});
});

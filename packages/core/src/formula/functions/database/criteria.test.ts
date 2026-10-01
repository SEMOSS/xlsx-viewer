import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { textToNumber } from '../../../values/parse-number';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { Scalar } from '../../types';
import { block } from '../lookup/test-utils/helpers';
import { criterionMatcher } from './criteria';
import { DCOUNT, DCOUNTA, DSUM } from './functions';
import { labelKey, readDatabase } from './table';

function matches(criterion: Scalar, values: Scalar[]): boolean[] {
	const m = criterionMatcher(criterion, false);
	return values.map((v) => (m ? m(v) : true));
}

describe('criterionMatcher', () => {
	it('treats blank criteria as no condition', () => {
		expect(criterionMatcher(null, false)).toBeNull();
		expect(criterionMatcher('', false)).toBeNull();
	});

	it('matches plain text as a case-insensitive prefix', () => {
		expect(matches('Dav', ['David', 'davis', 'Dave', 'Edavid', 'Da', 5])).toEqual([
			true,
			true,
			true,
			false,
			false,
			false,
		]);
	});

	it('matches =text exactly, with wildcards', () => {
		expect(matches('=Dav', ['Dav', 'DAV', 'David'])).toEqual([true, true, false]);
		expect(matches('=D*d', ['David', 'Dad', 'Davis'])).toEqual([true, true, false]);
		expect(matches('?a*', ['David', 'Bart', 'apple'])).toEqual([true, true, false]);
	});

	it('keeps literal wildcards and a trailing tilde', () => {
		expect(matches('a~*', ['a*b', 'abc'])).toEqual([true, false]);
		expect(matches('x~', ['x~1', 'x1'])).toEqual([true, false]);
		expect(matches('x~~', ['x~1', 'x1'])).toEqual([true, false]);
	});

	it('compares numbers, dates and text with operators', () => {
		expect(matches('>10', [5, 10, 11, '20', null])).toEqual([false, false, true, false, false]);
		expect(matches('<=10', [5, 10, 11])).toEqual([true, true, false]);
		expect(matches('<>10', [5, 10, 'x', null])).toEqual([true, false, true, true]);
		expect(matches('>M', ['apple', 'Nut', 'zoo', 5])).toEqual([false, true, true, false]);
		const day = textToNumber('3/15/2024') as number;
		expect(matches('>=3/15/2024', [day - 1, day, day + 1])).toEqual([false, true, true]);
	});

	it('matches blanks with "=" and non-blanks with "<>"', () => {
		expect(matches('=', [null, '', 0, 'x'])).toEqual([true, false, false, false]);
		expect(matches('<>', [null, '', 0, 'x'])).toEqual([false, false, true, true]);
	});

	it('matches numbers, logicals and errors by value', () => {
		expect(matches(12, [12, '12', 13, true])).toEqual([true, true, false, false]);
		expect(matches('12', [12, '12', '123'])).toEqual([true, true, false]);
		expect(matches(true, [true, 1, 'TRUE'])).toEqual([true, false, false]);
		expect(matches('FALSE', [false, true, 0])).toEqual([true, false, false]);
		expect(matches(ERRORS.NA, [ERRORS.NA, ERRORS.DIV0])).toEqual([true, false]);
		expect(matches('#DIV/0!', [ERRORS.DIV0, '#DIV/0!x'])).toEqual([true, false]);
	});
});

describe('database tables', () => {
	it('keys labels case-insensitively, first occurrence first', () => {
		expect(labelKey('Name')).toBe('name');
		expect(labelKey(2024)).toBe('2024');
		expect(labelKey(null)).toBeNull();
		expect(labelKey(ERRORS.NA)).toBe('#n/a');
		const db = readDatabase(arr(['A', 'b', 'a'], [1, 2, 3]), new MockContext());
		expect('columns' in db ? [...db.columns] : null).toEqual([
			['a', 0],
			['b', 1],
		]);
		expect('records' in db ? db.records : null).toBe(1);
	});
});

describe('criteria ranges', () => {
	// Database A1:C7; criteria blocks start at column E (index 4).
	const DATA = block([
		['Name', 'Dept', 'Salary'],
		['Ann', 'Sales', 50],
		['Bob', 'Sales', 70],
		['Cyd', 'IT', 90],
		['Dan', 'IT', 40],
		['Eve', 'HR', 60],
		['Ann-Marie', null, 80],
	]);
	const db = (c: MockContext) => c.ref('A1:C7');

	function sumWith(criteria: Scalar[][]): Scalar {
		const c = new MockContext({ ...DATA, ...block(criteria, 1, 4) });
		const rows = criteria.length;
		const cols = Math.max(...criteria.map((r) => r.length));
		const range = c.ref(`E1:${String.fromCharCode(69 + cols - 1)}${rows}`);
		return call(DSUM, [db(c), 'Salary', range], c) as Scalar;
	}

	it('ANDs conditions on one row', () => {
		expect(
			sumWith([
				['Dept', 'Salary'],
				['Sales', '>60'],
			]),
		).toBe(70);
		expect(
			sumWith([
				['Salary', 'Salary'],
				['>45', '<75'],
			]),
		).toBe(180);
	});

	it('ORs rows', () => {
		expect(sumWith([['Dept'], ['Sales'], ['HR']])).toBe(180);
		expect(
			sumWith([
				['Dept', 'Salary'],
				['IT', null],
				[null, '>75'],
			]),
		).toBe(210);
	});

	it('uses prefix matching for plain text and exact matching for =text', () => {
		expect(sumWith([['Name'], ['Ann']])).toBe(130);
		expect(sumWith([['Name'], ['=Ann']])).toBe(50);
		expect(sumWith([['name'], ['ann*']])).toBe(130);
	});

	it('matches blank fields with "=" and filled ones with "<>"', () => {
		expect(sumWith([['Dept'], ['=']])).toBe(80);
		expect(sumWith([['Dept'], ['<>']])).toBe(310);
	});

	it('matches every record for a row without conditions', () => {
		expect(sumWith([['Dept'], ['IT'], [null]])).toBe(390);
		expect(
			sumWith([
				['Dept', 'Name'],
				[null, null],
			]),
		).toBe(390);
	});

	it('extends to blank rows past the used range', () => {
		const c = new MockContext({ ...DATA, E1: 'Dept', E2: 'IT' });
		expect(call(DSUM, [db(c), 'Salary', c.ref('E1:E2')], c)).toBe(130);
		expect(call(DSUM, [db(c), 'Salary', c.ref('E1:E3')], c)).toBe(390);
		expect(call(DSUM, [db(c), 'Salary', c.ref('E:E')], c)).toBe(390);
	});

	it('never matches conditions under unknown or blank labels', () => {
		expect(sumWith([['Color'], ['red']])).toBe(0);
		expect(
			sumWith([
				['Color', 'Dept'],
				['red', 'IT'],
				[null, 'HR'],
			]),
		).toBe(60);
		expect(
			sumWith([
				[null, 'Dept'],
				[true, 'IT'],
			]),
		).toBe(0);
		expect(
			sumWith([
				['Color', 'Dept'],
				[null, 'IT'],
			]),
		).toBe(130);
	});

	it('counts with the same criteria rules', () => {
		const c = new MockContext({ ...DATA, ...block([['Name'], ['A']], 1, 4) });
		expect(call(DCOUNT, [db(c), 'Salary', c.ref('E1:E2')], c)).toBe(2);
		expect(call(DCOUNTA, [db(c), 'Dept', c.ref('E1:E2')], c)).toBe(1);
		expect(call(DCOUNTA, [db(c), null, c.ref('E1:E2')], c)).toBe(2);
	});
});

import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import {
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	parseRange,
	putCell,
	setCellInput,
} from 'xlsx-viewer-core';

import { findAll, makeMatcher, replaceInText, type FindHit, type FindOptions } from './find';

/**
 * Data: A1 'Apple pie', B1 'apple', A2 'Pineapple', B2 1234.5 shown as '$1,234.50',
 * A3 TRUE, B3 #N/A, C1 =B2+5 (1239.5). Other: A1 'apple tart', B1 'Banana'.
 */
function book(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Data', 'Other'] });
	const tx = new EditTransaction(wb, 'setup');
	const type = (sheet: number, ref: string, text: string) => {
		const { top, left } = parseRange(ref)!;
		setCellInput(tx, wb, sheet, top, left, text);
	};
	type(0, 'A1', 'Apple pie');
	type(0, 'B1', 'apple');
	type(0, 'A2', 'Pineapple');
	type(0, 'B2', '1234.5');
	applyStylePatch(tx, wb, 0, [parseRange('B2')!], { numFmt: '"$"#,##0.00' });
	type(0, 'A3', 'TRUE');
	type(0, 'B3', '#N/A');
	type(1, 'A1', 'apple tart');
	type(1, 'B1', 'Banana');
	tx.commit();
	putCell(wb.sheets[0], 0, 2, { value: 1239.5, formula: 'B2+5' });
	putCell(wb.sheets[0], 4, 0, { value: null, style: 1 });
	return wb;
}

const refs = (hits: FindHit[]) =>
	hits.map((h) => `${h.sheet}!${String.fromCharCode(65 + h.col)}${h.row + 1}`);

const find = (opts: FindOptions) => refs(findAll(book(), 0, opts));

describe('makeMatcher', () => {
	it('finds substrings, ignoring case unless asked', () => {
		expect(makeMatcher({ query: 'APP' })('Pineapple')).toBe(true);
		expect(makeMatcher({ query: 'pear' })('Pineapple')).toBe(false);
		const exact = makeMatcher({ query: 'App', matchCase: true });
		expect([exact('Apple'), exact('apple')]).toEqual([true, false]);
	});

	it('matches entire cell contents', () => {
		const whole = makeMatcher({ query: 'apple', wholeCell: true });
		expect([whole('Apple'), whole('apple pie'), whole('Pineapple')]).toEqual([true, false, false]);
		expect(makeMatcher({ query: 'apple', wholeCell: true, matchCase: true })('Apple')).toBe(false);
	});

	it('supports Excel wildcards: * for any run, ? for one character', () => {
		const star = makeMatcher({ query: 'a*e' });
		expect([star('Apple pie'), star('abc'), star('ae')]).toEqual([true, false, true]);
		const one = makeMatcher({ query: 'p?e' });
		expect([one('Apple pie'), one('pe'), one('ape')]).toEqual([true, false, false]);
		const wholeStar = makeMatcher({ query: 'a*', wholeCell: true });
		expect([wholeStar('apple'), wholeStar('a'), wholeStar('banana')]).toEqual([true, true, false]);
		expect(makeMatcher({ query: 'A*e', matchCase: true })('apple')).toBe(false);
	});

	it('escapes wildcards with a tilde', () => {
		const star = makeMatcher({ query: '5~*3' });
		expect([star('5*3'), star('513')]).toEqual([true, false]);
		const question = makeMatcher({ query: 'why~?', wholeCell: true });
		expect([question('why?'), question('whyx')]).toEqual([true, false]);
	});

	it('treats regular expression characters literally', () => {
		expect(makeMatcher({ query: 'a.b' })('axb')).toBe(false);
		expect(makeMatcher({ query: 'a.b' })('xa.by')).toBe(true);
		expect(makeMatcher({ query: '1+1', wholeCell: true })('1+1')).toBe(true);
		expect(makeMatcher({ query: '(x)*', wholeCell: true })('(x) and more')).toBe(true);
	});
});

describe('findAll', () => {
	it('finds nothing for an empty query', () => {
		expect(findAll(book(), 0, { query: '' })).toEqual([]);
	});

	it('searches the active sheet by rows, reporting the matched text', () => {
		const hits = findAll(book(), 0, { query: 'apple' });
		expect(hits).toEqual([
			{ sheet: 0, row: 0, col: 0, text: 'Apple pie' },
			{ sheet: 0, row: 0, col: 1, text: 'apple' },
			{ sheet: 0, row: 1, col: 0, text: 'Pineapple' },
		]);
		expect(refs(findAll(book(), 1, { query: 'apple' }))).toEqual(['1!A1']);
	});

	it('searches by columns', () => {
		expect(find({ query: 'apple', order: 'columns' })).toEqual(['0!A1', '0!A2', '0!B1']);
	});

	it('searches the whole workbook, sheet by sheet', () => {
		expect(find({ query: 'apple', within: 'workbook' })).toEqual(['0!A1', '0!B1', '0!A2', '1!A1']);
		expect(find({ query: 'an', within: 'workbook', order: 'columns' })).toEqual(['1!B1']);
	});

	it('honours match case and entire cell', () => {
		expect(find({ query: 'apple', matchCase: true, within: 'workbook' })).toEqual([
			'0!B1',
			'0!A2',
			'1!A1',
		]);
		expect(find({ query: 'APPLE', wholeCell: true })).toEqual(['0!B1']);
		expect(find({ query: 'APPLE', wholeCell: true, matchCase: true })).toEqual([]);
	});

	it('looks in formulas by default: formula text and unformatted numbers', () => {
		expect(find({ query: 'B2+' })).toEqual(['0!C1']);
		expect(find({ query: '1234.5' })).toEqual(['0!B2']);
		expect(find({ query: '1,234' })).toEqual([]);
		expect(find({ query: '1239' })).toEqual([]);
	});

	it('looks in displayed values when asked', () => {
		expect(find({ query: 'B2+', lookIn: 'values' })).toEqual([]);
		expect(find({ query: '$1,234.50', lookIn: 'values', wholeCell: true })).toEqual(['0!B2']);
		expect(find({ query: '1234.5', lookIn: 'values' })).toEqual([]);
		expect(find({ query: '1239', lookIn: 'values' })).toEqual(['0!C1']);
	});

	it('matches booleans and errors by their displayed text in both modes', () => {
		expect(find({ query: 'true' })).toEqual(['0!A3']);
		expect(find({ query: '#N/A', lookIn: 'values' })).toEqual(['0!B3']);
		expect(find({ query: '#n/a', wholeCell: true })).toEqual(['0!B3']);
	});

	it('never matches blank cells, even with a match-anything wildcard', () => {
		const all = find({ query: '*' });
		expect(all).toHaveLength(7);
		expect(all).not.toContain('0!A5');
	});
});

describe('replaceInText', () => {
	it('replaces every occurrence, ignoring case unless asked', () => {
		expect(replaceInText('Apple and apple', { query: 'apple' }, 'pear')).toBe('pear and pear');
		expect(replaceInText('Apple and apple', { query: 'apple', matchCase: true }, 'pear')).toBe(
			'Apple and pear',
		);
		expect(replaceInText('no match', { query: 'xyz' }, 'pear')).toBe('no match');
	});

	it('replaces the whole text for entire-cell matches', () => {
		expect(replaceInText('apple', { query: 'APPLE', wholeCell: true }, 'pear')).toBe('pear');
	});

	it('treats special characters in the query and the replacement literally', () => {
		expect(replaceInText('1+1=2, 1x1', { query: '1+1' }, 'two')).toBe('two=2, 1x1');
		expect(replaceInText('(a) a', { query: '(a)' }, 'b')).toBe('b a');
		expect(replaceInText('cost', { query: 'cost' }, '$&$1')).toBe('$&$1');
	});

	it('expands wildcards, matching as little as possible for *', () => {
		expect(replaceInText('abcabc', { query: 'a*c' }, 'X')).toBe('XX');
		expect(replaceInText('abc adc aXc', { query: 'a?c' }, 'Y')).toBe('Y Y Y');
		expect(replaceInText('ABC abc', { query: 'A?C', matchCase: true }, 'Z')).toBe('Z abc');
		expect(replaceInText('5*3 and 513', { query: '5~*3' }, 'n')).toBe('n and 513');
	});
});

import { describe, expect, it } from 'vitest';

import {
	escapeStructuredColumn,
	formatStructuredReference,
	parseStructuredReference,
} from './structured';

const parse = (text: string) => {
	const node = parseStructuredReference(text);
	if (node === null) {
		return null;
	}
	const { table, specifiers, columnStart, columnEnd } = node;
	return { table, specifiers, columnStart, columnEnd };
};

describe('parseStructuredReference', () => {
	it.each([
		['Table1[Sales]', { table: 'Table1', specifiers: [], columnStart: 'Sales' }],
		['Table1[Sales Amount]', { table: 'Table1', specifiers: [], columnStart: 'Sales Amount' }],
		['Table1[]', { table: 'Table1', specifiers: [] }],
		['Table1[#All]', { table: 'Table1', specifiers: ['#All'] }],
		['Table1[#this row]', { table: 'Table1', specifiers: ['#This Row'] }],
		[
			'Table1[[#Headers],[A]:[C]]',
			{ table: 'Table1', specifiers: ['#Headers'], columnStart: 'A', columnEnd: 'C' },
		],
		[
			'Table1[[#Headers],[#Data],[Q1]]',
			{ table: 'Table1', specifiers: ['#Headers', '#Data'], columnStart: 'Q1' },
		],
		[
			'Table1[ [Col1] : [Col2] ]',
			{ table: 'Table1', specifiers: [], columnStart: 'Col1', columnEnd: 'Col2' },
		],
		[
			'Table1[[Total $ Amount]]',
			{ table: 'Table1', specifiers: [], columnStart: 'Total $ Amount' },
		],
		['[@Qty]', { specifiers: ['#This Row'], columnStart: 'Qty' }],
		['[@]', { specifiers: ['#This Row'] }],
		[
			'Table1[@[Unit Price]]',
			{ table: 'Table1', specifiers: ['#This Row'], columnStart: 'Unit Price' },
		],
		['[@[A]:[B]]', { specifiers: ['#This Row'], columnStart: 'A', columnEnd: 'B' }],
		['[[#This Row],[Qty]]', { specifiers: ['#This Row'], columnStart: 'Qty' }],
		["Table1[Price'[USD']]", { table: 'Table1', specifiers: [], columnStart: 'Price[USD]' }],
		["Table1[['#Items]]", { table: 'Table1', specifiers: [], columnStart: '#Items' }],
		["Table1[It''s]", { table: 'Table1', specifiers: [], columnStart: "It's" }],
	])('%s', (text, expected) => {
		expect(parse(text)).toEqual({
			table: undefined,
			columnStart: undefined,
			columnEnd: undefined,
			...expected,
		});
	});

	it.each([
		'Table1[#Bogus]',
		'Table1[[#All],[#All]]',
		'Table1[[A],[B]]',
		'Table1[[A] x]',
		'Table1[[A]:]',
		'Table1',
	])('rejects %s', (text) => {
		expect(parseStructuredReference(text)).toBeNull();
	});

	it('keeps the source text', () => {
		expect(parseStructuredReference('T[ [A] ]')?.text).toBe('T[ [A] ]');
	});
});

describe('formatStructuredReference', () => {
	const node = (text: string) => parseStructuredReference(text)!;

	it.each([
		['Table1[Sales]', 'Table1[Sales]', 'Table1[Sales]'],
		['Table1[[Sales]]', 'Table1[Sales]', 'Table1[Sales]'],
		['Table1[#All]', 'Table1[#All]', 'Table1[#All]'],
		['Table1[]', 'Table1[]', 'Table1[]'],
		['[@Qty]', '[@Qty]', '[[#This Row],[Qty]]'],
		['[@]', '[@]', '[#This Row]'],
		['Table1[@[Unit-Price]]', 'Table1[@[Unit-Price]]', 'Table1[[#This Row],[Unit-Price]]'],
		['Table1[@[A]:[C]]', 'Table1[@[A]:[C]]', 'Table1[[#This Row],[A]:[C]]'],
		['Table1[[#Headers],[A]:[C]]', 'Table1[[#Headers],[A]:[C]]', 'Table1[[#Headers],[A]:[C]]'],
		["Table1[Price'[USD']]", "Table1[[Price'[USD']]]", "Table1[[Price'[USD']]]"],
		['Table1[[Total $]]', 'Table1[[Total $]]', 'Table1[[Total $]]'],
	])('%s -> %s / %s', (text, displayForm, fileForm) => {
		expect(formatStructuredReference(node(text))).toBe(displayForm);
		expect(formatStructuredReference(node(text), { forFile: true })).toBe(fileForm);
	});

	it('escapes special characters in column names', () => {
		expect(escapeStructuredColumn("a[b]#c'd")).toBe("a'[b']'#c''d");
	});
});

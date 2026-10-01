import { describe, expect, it } from 'vitest';

import { parseHtmlTable, parseTsv, toHtmlTable, toTsv } from './serialize';

describe('TSV', () => {
	it('round-trips quoted fields', () => {
		const rows = [
			['a', 'b\tc', 'line1\nline2'],
			['"q"', '', '3'],
		];
		expect(parseTsv(toTsv(rows))).toEqual(rows);
	});

	it('parses text from Excel', () => {
		expect(parseTsv('1\t2\r\n3\t4\r\n')).toEqual([
			['1', '2'],
			['3', '4'],
		]);
	});
});

describe('HTML tables', () => {
	it('expands colspans and reads cell text', () => {
		const html = toHtmlTable([[{ text: 'Title', colspan: 2 }], [{ text: '1' }, { text: '2' }]]);
		expect(parseHtmlTable(html)).toEqual([
			['Title', ''],
			['1', '2'],
		]);
	});
});

import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { csvToWorkbook, detectCsvDelimiter, parseCsvRows } from './csv-import';
import { sheetToCsv } from './index';

describe('CSV import', () => {
	it('parses quotes, doubled quotes and line breaks inside fields', () => {
		expect(parseCsvRows('a,"b,c","say ""hi""","two\nlines"\r\n1,2,3,4')).toEqual([
			['a', 'b,c', 'say "hi"', 'two\nlines'],
			['1', '2', '3', '4'],
		]);
		expect(parseCsvRows('\uFEFFx;y\n', ';')).toEqual([['x', 'y']]);
	});

	it('detects the delimiter', () => {
		expect(detectCsvDelimiter('a;b;c\n1;2;3\n')).toBe(';');
		expect(detectCsvDelimiter('a\tb\n1\t2\n')).toBe('\t');
		expect(detectCsvDelimiter('name,"x;y"\nA,B\n')).toBe(',');
	});

	it('recognises values like typing and never creates formulas', () => {
		const wb = csvToWorkbook(
			'Item,Price,Share,When,Ok,Formula\nPen,1.5,25%,2026-03-14,TRUE,=1+1\n',
			{ sheetName: 'prices' },
		);
		const sheet = wb.sheets[0];
		const v = (r: number, c: number) => sheet.cells.get(cellKey(r, c))?.value;
		expect(sheet.name).toBe('prices');
		expect(v(1, 1)).toBe(1.5);
		expect(v(1, 2)).toBe(0.25);
		expect(typeof v(1, 3)).toBe('number');
		expect(v(1, 4)).toBe(true);
		expect(v(1, 5)).toBe('=1+1');
		expect(sheet.cells.get(cellKey(1, 5))?.formula).toBeUndefined();
	});

	it('round-trips with sheetToCsv', () => {
		const text = 'a,b\n"x,y",2\n';
		const wb = csvToWorkbook(text);
		expect(sheetToCsv(wb, wb.sheets[0]).replace(/\r\n/g, '\n').trim()).toBe(text.trim());
	});
});

describe('readWorkbook with text files', () => {
	it('opens .csv and .tsv by file name', async () => {
		const { readWorkbook } = await import('../read');
		const csv = await readWorkbook(new TextEncoder().encode('a;b\n1;2\n'), {
			fileName: 'C:\\data\\Report 2026.csv',
		});
		expect(csv.sheets[0].name).toBe('Report 2026');
		expect(csv.sheets[0].cells.get(cellKey(1, 1))?.value).toBe(2);
		const tsv = await readWorkbook(new TextEncoder().encode('x,y\tz\n'), { fileName: 'list.tsv' });
		expect(tsv.sheets[0].cells.get(cellKey(0, 0))?.value).toBe('x,y');
		await expect(readWorkbook(new TextEncoder().encode('a,b'))).rejects.toThrow();
	});
});

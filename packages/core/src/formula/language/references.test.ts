import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import {
	formatReference,
	formatSheetQualifier,
	parseReferenceText,
	refAreaToRange,
	sameSheetName,
} from './references';

describe('formatSheetQualifier', () => {
	it.each([
		[{ sheet: 'Sheet1' }, 'Sheet1!'],
		[{ sheet: 'My Sheet' }, "'My Sheet'!"],
		[{ sheet: "Bob's" }, "'Bob''s'!"],
		[{ sheet: 'Sheet1', sheetEnd: 'Sheet3' }, 'Sheet1:Sheet3!'],
		[{ sheet: 'Sheet 1', sheetEnd: 'Sheet3' }, "'Sheet 1:Sheet3'!"],
		[{ sheet: 'Sheet1', workbook: '1' }, '[1]Sheet1!'],
		[{ sheet: 'Sheet 1', workbook: '1' }, "'[1]Sheet 1'!"],
		[{ sheet: 'Data', workbook: 'Book 1.xlsx' }, "'[Book 1.xlsx]Data'!"],
		[{ sheet: '', workbook: '1' }, '[1]!'],
		[{ sheet: '', workbook: 'C:\\x\\Book.xlsx' }, "'C:\\x\\Book.xlsx'!"],
	])('%j -> %s', (q, text) => {
		expect(formatSheetQualifier(q)).toBe(text);
	});
});

describe('parseReferenceText', () => {
	it('parses a single reference', () => {
		expect(parseReferenceText('Sheet1!$A$1:$B$2')).toEqual({
			sheet: { sheet: 'Sheet1' },
			area: {
				kind: 'cells',
				start: { row: 0, col: 0, rowAbs: true, colAbs: true },
				end: { row: 1, col: 1, rowAbs: true, colAbs: true },
			},
		});
		expect(parseReferenceText(" 'My Sheet'!C:C ")?.area.kind).toBe('columns');
		expect(parseReferenceText('=B2')?.area.start).toEqual({
			row: 1,
			col: 1,
			rowAbs: false,
			colAbs: false,
		});
		expect(parseReferenceText('R[1]C', { r1c1: { row: 0, col: 0 } })?.area.start.row).toBe(1);
	});

	it.each(['', 'A1+B1', 'Rate', '"A1"', 'A1 B1', 'SUM(A1)', 'Sheet1!', '#REF!'])(
		'rejects %j',
		(text) => {
			expect(parseReferenceText(text)).toBeNull();
		},
	);

	it('round-trips with formatReference', () => {
		for (const text of ["'Q1 Data'!$A$1:B2", 'Sheet1:Sheet2!C3', '[1]Sheet1!$1:$2', 'XFD1048576']) {
			const parsed = parseReferenceText(text)!;
			expect(formatReference(parsed.sheet, parsed.area)).toBe(text);
		}
	});
});

describe('reference helpers', () => {
	it('normalizes areas to ranges', () => {
		expect(refAreaToRange(parseReferenceText('C5:A1')!.area)).toEqual({
			top: 0,
			left: 0,
			bottom: 4,
			right: 2,
		});
		expect(refAreaToRange(parseReferenceText('2:4')!.area)).toEqual({
			top: 1,
			left: 0,
			bottom: 3,
			right: MAX_COL_INDEX,
		});
		expect(refAreaToRange(parseReferenceText('B:B')!.area)).toEqual({
			top: 0,
			left: 1,
			bottom: MAX_ROW_INDEX,
			right: 1,
		});
	});

	it('compares sheet names case-insensitively', () => {
		expect(sameSheetName('Sheet1', 'SHEET1')).toBe(true);
		expect(sameSheetName('Données', 'DONNÉES')).toBe(true);
		expect(sameSheetName('Sheet1', 'Sheet2')).toBe(false);
	});
});

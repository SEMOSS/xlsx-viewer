import { describe, expect, it } from 'vitest';

import {
	cellKey,
	columnToLetters,
	formatRange,
	keyCol,
	keyRow,
	lettersToColumn,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	parseCellAddress,
	parseRange,
	parseSqref,
	quoteSheetName,
	unquoteSheetName,
} from './index';

describe('column letters', () => {
	it.each([
		[0, 'A'],
		[25, 'Z'],
		[26, 'AA'],
		[701, 'ZZ'],
		[702, 'AAA'],
		[16383, 'XFD'],
	])('%i <-> %s', (col, letters) => {
		expect(columnToLetters(col)).toBe(letters);
		expect(lettersToColumn(letters)).toBe(col);
	});

	it('rejects invalid letters', () => {
		expect(lettersToColumn('XFE')).toBe(-1);
		expect(lettersToColumn('')).toBe(-1);
		expect(lettersToColumn('A1')).toBe(-1);
	});
});

describe('cell keys', () => {
	it('round-trips and sorts in row-major order', () => {
		const k = cellKey(1048575, 16383);
		expect(keyRow(k)).toBe(1048575);
		expect(keyCol(k)).toBe(16383);
		expect(cellKey(0, 5) < cellKey(1, 0)).toBe(true);
	});
});

describe('A1 parsing', () => {
	it('parses cells and ranges', () => {
		expect(parseCellAddress('$B$3')).toEqual({ row: 2, col: 1 });
		expect(parseRange('B3:A1')).toEqual({ top: 0, left: 0, bottom: 2, right: 1 });
		expect(parseRange('C:A')).toEqual({ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: 2 });
		expect(parseRange('$2:$4')).toEqual({ top: 1, left: 0, bottom: 3, right: MAX_COL_INDEX });
		expect(parseCellAddress('A0')).toBeNull();
		expect(parseCellAddress('A1048577')).toBeNull();
	});

	it('formats ranges', () => {
		expect(formatRange({ top: 0, left: 0, bottom: 0, right: 0 })).toBe('A1');
		expect(formatRange({ top: 0, left: 0, bottom: 9, right: 2 }, { absolute: true })).toBe(
			'$A$1:$C$10',
		);
		expect(formatRange({ top: 0, left: 1, bottom: MAX_ROW_INDEX, right: 1 })).toBe('B:B');
		expect(formatRange({ top: 4, left: 0, bottom: 4, right: MAX_COL_INDEX })).toBe('5:5');
	});

	it('parses sqref lists', () => {
		expect(parseSqref('A1:B2 D4')).toHaveLength(2);
	});
});

describe('sheet names', () => {
	it('quotes when needed', () => {
		expect(quoteSheetName('Sheet1')).toBe('Sheet1');
		expect(quoteSheetName('My Sheet')).toBe("'My Sheet'");
		expect(quoteSheetName("Bob's")).toBe("'Bob''s'");
		expect(quoteSheetName('A1')).toBe("'A1'");
		expect(quoteSheetName('2024')).toBe("'2024'");
		expect(unquoteSheetName("'Bob''s'")).toBe("Bob's");
	});
});

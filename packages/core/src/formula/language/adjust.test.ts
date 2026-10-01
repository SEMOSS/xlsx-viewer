import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import { parseRange } from '../../address/range';
import type { RangeAddress } from '../../types/address';
import { adjustFormulaForStructureChange } from './adjust';
import { adjustRangeForStructureChange, type StructureChange } from './adjust-range';

const insertRows = (at: number, count = 1, sheet = 'Sheet1'): StructureChange => ({
	kind: 'insertRows',
	sheet,
	at,
	count,
});
const deleteRows = (at: number, count = 1, sheet = 'Sheet1'): StructureChange => ({
	kind: 'deleteRows',
	sheet,
	at,
	count,
});
const insertCols = (at: number, count = 1, sheet = 'Sheet1'): StructureChange => ({
	kind: 'insertCols',
	sheet,
	at,
	count,
});
const deleteCols = (at: number, count = 1, sheet = 'Sheet1'): StructureChange => ({
	kind: 'deleteCols',
	sheet,
	at,
	count,
});

const adjust = (formula: string, change: StructureChange, formulaSheet = 'Sheet1') =>
	adjustFormulaForStructureChange(formula, change, formulaSheet).formula;

describe('adjustFormulaForStructureChange: inserting rows', () => {
	// Insert one row above Excel row 2 (0-based index 1).
	it.each([
		['A1', 'A1'],
		['A2', 'A3'],
		['$A$2', '$A$3'],
		['A$2+$A2', 'A$3+$A3'],
		['SUM(A1:A3)', 'SUM(A1:A4)'],
		['SUM(A2:A3)', 'SUM(A3:A4)'],
		['SUM(A1:A1)', 'SUM(A1:A1)'],
		['SUM(A5:B10)', 'SUM(A6:B11)'],
		['SUM($A$1:$A$3)', 'SUM($A$1:$A$4)'],
		['SUM(A3:A1)', 'SUM(A4:A1)'],
		['SUM(A:A)', 'SUM(A:A)'],
		['SUM(2:5)', 'SUM(3:6)'],
		['SUM(1:3)', 'SUM(1:4)'],
		['SUM(1:1)', 'SUM(1:1)'],
		['Sheet1!A2', 'Sheet1!A3'],
		["'sheet1'!A2", "'sheet1'!A3"],
		['Sheet2!A2', 'Sheet2!A2'],
		['A2#', 'A3#'],
		['"A2"&A2&Rate&Table1[A2]', '"A2"&A3&Rate&Table1[A2]'],
		['SUM( A2 ,\n A5 )', 'SUM( A3 ,\n A6 )'],
	])('%s -> %s', (formula, expected) => {
		expect(adjust(formula, insertRows(1))).toBe(expected);
	});

	it('expands a range when inserting at its last row but not just below it', () => {
		expect(adjust('SUM(A1:A3)', insertRows(2))).toBe('SUM(A1:A4)');
		expect(adjust('SUM(A1:A3)', insertRows(3))).toBe('SUM(A1:A3)');
		expect(adjust('SUM(A2:A4)', insertRows(0, 3))).toBe('SUM(A5:A7)');
	});

	it('only touches references to the changed sheet', () => {
		expect(adjust('A2+Sheet1!A2', insertRows(1, 1, 'Sheet1'), 'Sheet2')).toBe('A2+Sheet1!A3');
		expect(adjust('Sheet1:Sheet3!A2+[1]Sheet1!A2', insertRows(1))).toBe(
			'Sheet1:Sheet3!A2+[1]Sheet1!A2',
		);
	});

	it('treats full-height areas like whole columns', () => {
		expect(adjust('SUM(A1:A1048576)', insertRows(0))).toBe('SUM(A1:A1048576)');
	});

	it('drops references pushed off the grid', () => {
		expect(adjust('A1048576+A1', insertRows(0))).toBe('#REF!+A2');
		expect(adjust('SUM(A5:A1048576)', insertRows(0))).toBe('SUM(A6:A1048576)');
		expect(adjust('SUM(A1048575:A1048576)', insertRows(0, 5))).toBe('SUM(#REF!)');
	});

	it('reports whether anything changed', () => {
		expect(adjustFormulaForStructureChange('A1+B1', insertRows(5), 'Sheet1')).toEqual({
			formula: 'A1+B1',
			changed: false,
		});
		expect(adjustFormulaForStructureChange('A9', insertRows(5), 'Sheet1')).toEqual({
			formula: 'A10',
			changed: true,
		});
		expect(adjustFormulaForStructureChange('A9', insertRows(5, 0), 'Sheet1').changed).toBe(false);
	});
});

describe('adjustFormulaForStructureChange: deleting rows', () => {
	// Delete Excel rows 2-3 (0-based 1, count 2).
	it.each([
		['A1', 'A1'],
		['A2', '#REF!'],
		['A3+1', '#REF!+1'],
		['A4', 'A2'],
		['$A$4', '$A$2'],
		['SUM(A1:A5)', 'SUM(A1:A3)'],
		['SUM(A2:A3)', 'SUM(#REF!)'],
		['SUM(A2:A5)', 'SUM(A2:A3)'],
		['SUM(A1:A2)', 'SUM(A1:A1)'],
		['SUM(A3:A1)', 'SUM(A1:A1)'],
		['SUM(A2:A3,A4)', 'SUM(#REF!,A2)'],
		['Sheet1!A2', 'Sheet1!#REF!'],
		["'Sheet1'!B3:C9", "'Sheet1'!B2:C7"],
		['SUM(2:3)', 'SUM(#REF!)'],
		['SUM(1:10)', 'SUM(1:8)'],
		['SUM(B:B)', 'SUM(B:B)'],
		['A2#+1', '#REF!+1'],
		['A5#', 'A3#'],
	])('%s -> %s', (formula, expected) => {
		expect(adjust(formula, deleteRows(1, 2))).toBe(expected);
	});
});

describe('adjustFormulaForStructureChange: columns', () => {
	it('inserts columns', () => {
		expect(adjust('B1+A1', insertCols(1))).toBe('C1+A1');
		expect(adjust('SUM(A1:C1)', insertCols(1))).toBe('SUM(A1:D1)');
		expect(adjust('SUM(B:C)', insertCols(1))).toBe('SUM(C:D)');
		expect(adjust('SUM(A:A)', insertCols(1))).toBe('SUM(A:A)');
		expect(adjust('SUM(1:1)', insertCols(0))).toBe('SUM(1:1)');
		expect(adjust('SUM(A1:XFD1)', insertCols(0))).toBe('SUM(A1:XFD1)');
		expect(adjust('XFD1', insertCols(0))).toBe('#REF!');
	});

	it('deletes columns', () => {
		expect(adjust('B1', deleteCols(1))).toBe('#REF!');
		expect(adjust('C1', deleteCols(1))).toBe('B1');
		expect(adjust('SUM(A1:C1)', deleteCols(1))).toBe('SUM(A1:B1)');
		expect(adjust('SUM($B$1:$D$9)', deleteCols(0, 2))).toBe('SUM($A$1:$B$9)');
		expect(adjust('SUM(B:B)', deleteCols(1))).toBe('SUM(#REF!)');
		expect(adjust('SUM(2:2)', deleteCols(1))).toBe('SUM(2:2)');
	});
});

describe('adjustRangeForStructureChange', () => {
	const r = (text: string) => parseRange(text) as RangeAddress;

	it('shifts, expands, shrinks and deletes ranges', () => {
		expect(adjustRangeForStructureChange(r('A2:B4'), insertRows(1))).toEqual(r('A3:B5'));
		expect(adjustRangeForStructureChange(r('A2:B4'), insertRows(2, 2))).toEqual(r('A2:B6'));
		expect(adjustRangeForStructureChange(r('A2:B4'), deleteRows(2))).toEqual(r('A2:B3'));
		expect(adjustRangeForStructureChange(r('A2:B4'), deleteRows(0, 10))).toBeNull();
		expect(adjustRangeForStructureChange(r('B1:D1'), insertCols(2))).toEqual(r('B1:E1'));
		expect(adjustRangeForStructureChange(r('B1:D1'), deleteCols(0))).toEqual(r('A1:C1'));
		expect(adjustRangeForStructureChange(r('B1:B1'), deleteCols(1))).toBeNull();
	});

	it('returns the same object when nothing changes', () => {
		const range = r('A1:B2');
		expect(adjustRangeForStructureChange(range, insertRows(5))).toBe(range);
		expect(adjustRangeForStructureChange(range, insertCols(5))).toBe(range);
		expect(adjustRangeForStructureChange(range, deleteRows(5, 0))).toBe(range);
	});

	it('leaves whole columns alone for row changes and whole rows for column changes', () => {
		const column: RangeAddress = { top: 0, left: 1, bottom: MAX_ROW_INDEX, right: 1 };
		expect(adjustRangeForStructureChange(column, insertRows(0))).toBe(column);
		expect(adjustRangeForStructureChange(column, deleteRows(0, 100))).toBe(column);
		expect(adjustRangeForStructureChange(column, insertCols(0))).toEqual({
			...column,
			left: 2,
			right: 2,
		});
		const row: RangeAddress = { top: 3, left: 0, bottom: 3, right: MAX_COL_INDEX };
		expect(adjustRangeForStructureChange(row, insertCols(0))).toBe(row);
		expect(adjustRangeForStructureChange(row, deleteRows(3))).toBeNull();
	});

	it('clamps ranges pushed past the last row', () => {
		expect(
			adjustRangeForStructureChange(
				{ top: 10, left: 0, bottom: MAX_ROW_INDEX, right: 0 },
				insertRows(5),
			),
		).toEqual({
			top: 11,
			left: 0,
			bottom: MAX_ROW_INDEX,
			right: 0,
		});
		expect(
			adjustRangeForStructureChange(
				{ top: MAX_ROW_INDEX, left: 0, bottom: MAX_ROW_INDEX, right: 0 },
				insertRows(0),
			),
		).toBeNull();
	});
});

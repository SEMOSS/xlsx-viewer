import { describe, expect, it } from 'vitest';

import { MAX_ROW_INDEX } from '../address/constants';
import { parseRange } from '../address/range';
import {
	adjustFormulaForBand,
	adjustRangeForBand,
	bandFateOf,
	bandShiftOf,
	bandZones,
	movePointForBand,
	structureChangeOf,
} from './shift-band';

const r = (a1: string) => parseRange(a1)!;

describe('bandShiftOf / structureChangeOf', () => {
	it('describes vertical and horizontal shifts', () => {
		expect(bandShiftOf('S', r('B3:C4'), 'down')).toEqual({
			sheet: 'S',
			vertical: true,
			insert: true,
			at: 2,
			count: 2,
			from: 1,
			to: 2,
		});
		expect(bandShiftOf('S', r('B3:D4'), 'left')).toEqual({
			sheet: 'S',
			vertical: false,
			insert: false,
			at: 1,
			count: 3,
			from: 2,
			to: 3,
		});
		expect(structureChangeOf(bandShiftOf('S', r('B3'), 'up'))).toEqual({
			kind: 'deleteRows',
			sheet: 'S',
			at: 2,
			count: 1,
		});
		expect(structureChangeOf(bandShiftOf('S', r('B3'), 'right')).kind).toBe('insertCols');
	});
});

describe('adjustFormulaForBand', () => {
	const down = bandShiftOf('Sheet1', r('B3'), 'down');
	const up = bandShiftOf('Sheet1', r('B3'), 'up');
	const right = bandShiftOf('Sheet1', r('B2'), 'right');
	const fix = (f: string, band = down, sheet = 'Sheet1') => adjustFormulaForBand(f, band, sheet);

	it.each([
		['B5', 'B6'],
		['$B$5', '$B$6'],
		['SUM(B1:B5)', 'SUM(B1:B6)'],
		['SUM(B3:B5)', 'SUM(B4:B6)'],
		['SUM(A1:C5)', 'SUM(A1:C5)'],
		['B2', 'B2'],
		['A5+C5', 'A5+C5'],
		['SUM(B:B)', 'SUM(B:B)'],
		['SUM(3:3)', 'SUM(3:3)'],
		['"B5"&B5', '"B5"&B6'],
	])('insert B3 shift down: %s -> %s', (formula, expected) => {
		expect(fix(formula)).toBe(expected);
	});

	it.each([
		['B3', '#REF!'],
		['SUM(B1:B5)', 'SUM(B1:B4)'],
		['B6', 'B5'],
		['B2', 'B2'],
		['SUM(B3:B3)', 'SUM(#REF!)'],
		['C6', 'C6'],
	])('delete B3 shift up: %s -> %s', (formula, expected) => {
		expect(fix(formula, up)).toBe(expected);
	});

	it.each([
		['C2', 'D2'],
		['SUM(A2:D2)', 'SUM(A2:E2)'],
		['SUM(A1:D3)', 'SUM(A1:D3)'],
		['SUM(2:2)', 'SUM(2:2)'],
		['SUM(C:C)', 'SUM(C:C)'],
		['A2', 'A2'],
	])('insert B2 shift right: %s -> %s', (formula, expected) => {
		expect(fix(formula, right)).toBe(expected);
	});

	it('follows sheet qualifiers', () => {
		expect(fix('Sheet1!B5+B5', down, 'Sheet2')).toBe('Sheet1!B6+B5');
		expect(fix("'sheet1'!B3", up, 'Other')).toBe("'sheet1'!#REF!");
		expect(fix('Sheet1:Sheet3!B5', down, 'Sheet2')).toBe('Sheet1:Sheet3!B5');
		expect(fix('Sheet2!B5+B5', down, 'Sheet1')).toBe('Sheet2!B5+B6');
		expect(fix("'My Data'!B5", { ...down, sheet: 'My Data' }, 'Sheet1')).toBe("'My Data'!B6");
	});

	it('moves spill references and drops the spill operator of deleted ones', () => {
		expect(fix('SUM(B5#)')).toBe('SUM(B6#)');
		expect(fix('SUM(B3#)+1', up)).toBe('SUM(#REF!)+1');
	});

	it('treats full-height areas like whole columns for vertical shifts', () => {
		expect(fix(`SUM(B1:B${MAX_ROW_INDEX + 1})`)).toBe(`SUM(B1:B${MAX_ROW_INDEX + 1})`);
	});
});

describe('adjustRangeForBand / movePointForBand', () => {
	const down = bandShiftOf('S', r('B3:C3'), 'down');
	const up = bandShiftOf('S', r('B3:C4'), 'up');

	it('adjusts ranges inside the band cross extent only', () => {
		expect(adjustRangeForBand(r('B5:C6'), down)).toEqual(r('B6:C7'));
		expect(adjustRangeForBand(r('B1:B5'), down)).toEqual(r('B1:B6'));
		const wide = r('A5:C6');
		expect(adjustRangeForBand(wide, down)).toBe(wide);
		expect(adjustRangeForBand(r('B3:C4'), up)).toBeNull();
		expect(adjustRangeForBand(r('B1:B9'), up)).toEqual(r('B1:B7'));
	});

	it('moves, keeps or deletes points', () => {
		const p = { row: 1, col: 1 };
		expect(movePointForBand(p, down)).toBe(p);
		expect(movePointForBand({ row: 2, col: 2 }, down)).toEqual({ row: 3, col: 2 });
		expect(movePointForBand({ row: 2, col: 3 }, down)).toEqual({ row: 2, col: 3 });
		expect(movePointForBand({ row: 3, col: 1 }, up)).toBeNull();
		expect(movePointForBand({ row: 6, col: 1, extra: 'x' }, up)).toEqual({
			row: 4,
			col: 1,
			extra: 'x',
		});
		expect(movePointForBand({ row: MAX_ROW_INDEX, col: 1 }, down)).toBeNull();
	});
});

describe('bandZones / bandFateOf', () => {
	it('classifies objects against the moving block and the deleted cells', () => {
		const ins = bandZones(bandShiftOf('S', r('B3:C3'), 'down'));
		expect(bandFateOf(r('A1:D2'), ins)).toBe('outside');
		expect(bandFateOf(r('B5:C6'), ins)).toBe('moves');
		expect(bandFateOf(r('B2:B3'), ins)).toBe('split');
		expect(bandFateOf(r('A5:B5'), ins)).toBe('split');
		const del = bandZones(bandShiftOf('S', r('B3:C4'), 'up'));
		expect(bandFateOf(r('B3:C4'), del)).toBe('removed');
		expect(bandFateOf(r('B5:C5'), del)).toBe('moves');
		expect(bandFateOf(r('B4:B5'), del)).toBe('split');
		expect(bandFateOf(r('D3:D4'), del)).toBe('outside');
	});
});

import { describe, expect, it } from 'vitest';
import type { FillSpec, RangeAddress, Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	applyStylePatch,
	cellKey,
	createWorkbook,
	EditTransaction,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';

import { MergeIndex, styleIdAt } from './style-at';

const r = (ref: string): RangeAddress => parseRange(ref)!;
const solid = (rgb: string): FillSpec => ({
	type: 'pattern',
	patternType: 'solid',
	fgColor: { rgb },
});

function sheetWith(build: (tx: EditTransaction, wb: Workbook) => void): Worksheet {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	build(tx, wb);
	tx.commit();
	return wb.sheets[0];
}

describe('styleIdAt', () => {
	// Column C and row 3 are formatted; C3 sits at their intersection.
	const sheet = sheetWith((tx, wb) => {
		applyStylePatch(tx, wb, 0, [r('C:C')], { fill: solid('FFFF0000') });
		applyStylePatch(tx, wb, 0, [r('3:3')], { fill: solid('FF00FF00') });
		setCellInput(tx, wb, 0, 0, 0, 'plain');
		setCellInput(tx, wb, 0, 4, 4, 'styled');
		applyStylePatch(tx, wb, 0, [r('E5')], { font: { bold: true } });
	});
	const colStyle = sheet.cols.find((c) => c.min <= 2 && c.max >= 2)?.style;
	const rowStyle = sheet.rows.get(2)?.style;

	it('uses the cell style first', () => {
		const own = sheet.cells.get(cellKey(4, 4))?.style;
		expect(own).toBeGreaterThan(0);
		expect(styleIdAt(sheet, 4, 4)).toBe(own);
	});

	it('falls back to the row style, then the column style, then Normal', () => {
		expect(colStyle).toBeGreaterThan(0);
		expect(rowStyle).toBeGreaterThan(0);
		expect(styleIdAt(sheet, 2, 7)).toBe(rowStyle);
		expect(styleIdAt(sheet, 9, 2)).toBe(colStyle);
		expect(styleIdAt(sheet, 9, 7)).toBe(0);
	});

	it('lets cells without their own style inherit the row or column format', () => {
		// A1 was typed without formatting: it has no style of its own.
		expect(sheet.cells.get(cellKey(0, 0))?.style).toBeUndefined();
		expect(styleIdAt(sheet, 0, 0)).toBe(0);
		const typed = sheetWith((tx, wb) => {
			applyStylePatch(tx, wb, 0, [r('C:C')], { fill: solid('FFFF0000') });
			setCellInput(tx, wb, 0, 5, 2, 'in C');
		});
		expect(styleIdAt(typed, 5, 2)).toBe(typed.cols[0].style);
	});

	it('ignores row styles without customFormat', () => {
		const s = sheetWith(() => undefined);
		s.rows.set(1, { style: 7 });
		s.cols = [{ min: 0, max: 0, style: 3 }];
		expect(styleIdAt(s, 1, 0)).toBe(3);
		s.rows.set(1, { style: 7, customFormat: true });
		expect(styleIdAt(s, 1, 0)).toBe(7);
	});

	it('uses the cell passed in instead of looking it up', () => {
		expect(styleIdAt(sheet, 9, 9, { value: 1, style: 5 })).toBe(5);
		expect(styleIdAt(sheet, 2, 9, { value: 1 })).toBe(rowStyle);
	});
});

describe('MergeIndex', () => {
	const merges = [r('B2:C3'), r('E1:E4'), r('A10:B10')];
	const index = new MergeIndex(merges);

	it('finds the merge containing any of its cells', () => {
		for (const [row, col] of [
			[1, 1],
			[1, 2],
			[2, 1],
			[2, 2],
		]) {
			expect(index.find(row, col)).toBe(merges[0]);
		}
		expect(index.find(3, 4)).toBe(merges[1]);
		expect(index.find(9, 1)).toBe(merges[2]);
	});

	it('returns nothing outside merges', () => {
		expect(index.find(0, 0)).toBeUndefined();
		expect(index.find(1, 3)).toBeUndefined();
		expect(index.find(4, 4)).toBeUndefined();
		expect(new MergeIndex([]).find(0, 0)).toBeUndefined();
	});

	it('resolves huge merges by scanning', () => {
		const huge = r('A1:Z2000');
		const big = new MergeIndex([huge, r('AB1:AC2')]);
		expect(big.find(0, 0)).toBe(huge);
		expect(big.find(1999, 25)).toBe(huge);
		expect(big.find(2000, 0)).toBeUndefined();
		expect(big.find(0, 26)).toBeUndefined();
		expect(big.find(1, 28)).toEqual(r('AB1:AC2'));
	});

	it('exposes the merges it was built from', () => {
		expect(index.merges).toBe(merges);
	});
});

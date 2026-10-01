import { describe, expect, it } from 'vitest';
import type { Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	columnWidthToPixels,
	createWorkbook,
	createWorksheet,
	EditTransaction,
	MAX_COL_INDEX,
	setColumnsHidden,
	setColumnWidth,
	setRowHeight,
	setRowsHidden,
} from 'xlsx-viewer-core';

import { SheetGeometry } from './sheet-geometry';

function edit(build: (tx: EditTransaction, wb: Workbook) => void): Worksheet {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	build(tx, wb);
	tx.commit();
	return wb.sheets[0];
}

describe('SheetGeometry: column widths', () => {
	it('converts stored widths to pixels with the Excel formula', () => {
		const widths = [1, 2.5, 8.43, 9.140625, 10.7109375, 13, 20, 50, 255];
		const sheet = edit((tx, wb) => {
			widths.forEach((w, i) => setColumnWidth(tx, wb, 0, i, i, w));
		});
		const g = new SheetGeometry(sheet, { zoom: 1 });
		widths.forEach((w, i) => expect(g.colWidth(i), `width ${w}`).toBe(columnWidthToPixels(w, 7)));
		// Known values: 1 char = 7 px, Excel's default 8.43 chars = 64 px, 10 chars = 75 px.
		expect(widths.map((_, i) => g.colWidth(i))).toEqual([7, 17, 59, 64, 75, 91, 140, 350, 1785]);
	});

	it('uses the maximum digit width of the workbook font', () => {
		const sheet = createWorksheet('S', 1);
		sheet.cols = [{ min: 1, max: 1, width: 20 }];
		const g = new SheetGeometry(sheet, { zoom: 1, maxDigitWidth: 8 });
		expect(g.defaultColWidth).toBe(73);
		expect(g.colWidth(0)).toBe(73);
		expect(g.colWidth(1)).toBe(160);
	});

	it('takes the default width from defaultColWidth, else baseColWidth', () => {
		const sheet = createWorksheet('S', 1);
		expect(new SheetGeometry(sheet, { zoom: 1 }).defaultColWidth).toBe(64);
		sheet.baseColWidth = 10;
		// 10 digits + 5 px padding, rounded up to a multiple of 8 px.
		expect(new SheetGeometry(sheet, { zoom: 1 }).defaultColWidth).toBe(80);
		sheet.defaultColWidth = 12;
		expect(new SheetGeometry(sheet, { zoom: 1 }).colWidth(5)).toBe(84);
	});

	it('applies column records over their whole span and clamps them to the sheet', () => {
		const sheet = createWorksheet('S', 1);
		sheet.cols = [
			{ min: 2, max: 4, width: 20 },
			{ min: 16000, max: 20000, width: 1 },
		];
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect([1, 2, 3, 4, 5].map((c) => g.colWidth(c))).toEqual([64, 140, 140, 140, 64]);
		expect(g.colWidth(MAX_COL_INDEX)).toBe(7);
		expect(g.colLeft(5)).toBe(2 * 64 + 3 * 140);
	});

	it('gives hidden and zero-width columns no width', () => {
		const sheet = edit((tx, wb) => {
			setColumnsHidden(tx, wb, 0, 1, 2, true);
			setColumnWidth(tx, wb, 0, 4, 4, 0);
		});
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect([0, 1, 2, 3, 4, 5].map((c) => g.colWidth(c))).toEqual([64, 0, 0, 64, 0, 64]);
		expect(g.colLeft(3)).toBe(64);
		expect(g.colLeft(5)).toBe(128);
		expect(g.cols.isHidden(2)).toBe(true);
	});
});

describe('SheetGeometry: row heights', () => {
	it('converts point heights to rounded pixels', () => {
		const sheet = edit((tx, wb) => {
			setRowHeight(tx, wb, 0, 1, 1, 30);
			setRowHeight(tx, wb, 0, 2, 2, 14.4);
			setRowHeight(tx, wb, 0, 3, 3, 12.75);
		});
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect([0, 1, 2, 3, 4].map((r) => g.rowHeight(r))).toEqual([20, 40, 19, 17, 20]);
		expect(g.rowTop(4)).toBe(20 + 40 + 19 + 17);
	});

	it('follows the sheet default row height', () => {
		const sheet = createWorksheet('S', 1);
		sheet.defaultRowHeight = 18;
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect(g.defaultRowHeight).toBe(24);
		expect(g.rowTop(10)).toBe(240);
	});

	it('gives hidden rows no height', () => {
		const sheet = edit((tx, wb) => {
			setRowsHidden(tx, wb, 0, 2, 3, true);
			setRowHeight(tx, wb, 0, 5, 5, 0);
		});
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect([1, 2, 3, 4, 5, 6].map((r) => g.rowHeight(r))).toEqual([20, 0, 0, 20, 0, 20]);
		expect(g.rowTop(4)).toBe(40);
		expect(g.rows.isHidden(3)).toBe(true);
	});

	it('hides unlisted rows when rows are hidden by default (zeroHeight)', () => {
		const sheet = createWorksheet('S', 1);
		sheet.zeroHeight = true;
		sheet.rows.set(1, { customFormat: true, style: 0 });
		sheet.rows.set(2, { height: 30 });
		sheet.rows.set(3, { hidden: true, height: 30 });
		const g = new SheetGeometry(sheet, { zoom: 1 });
		expect(g.defaultRowHeight).toBe(0);
		expect([0, 1, 2, 3, 4].map((r) => g.rowHeight(r))).toEqual([0, 20, 40, 0, 0]);
	});
});

describe('SheetGeometry: zoom', () => {
	const sheet = edit((tx, wb) => {
		setColumnWidth(tx, wb, 0, 1, 1, 20);
		setRowHeight(tx, wb, 0, 1, 1, 30);
		setColumnsHidden(tx, wb, 0, 2, 2, true);
	});

	it('scales and rounds every line', () => {
		const g = new SheetGeometry(sheet, { zoom: 1.5 });
		expect(g.zoom).toBe(1.5);
		expect([g.defaultColWidth, g.defaultRowHeight]).toEqual([96, 30]);
		expect([g.colWidth(1), g.rowHeight(1), g.colWidth(2)]).toEqual([210, 60, 0]);
		expect(g.colLeft(4)).toBe(96 + 210 + 0 + 96);
	});

	it('keeps lines at least 1 px when zoomed far out', () => {
		const g = new SheetGeometry(sheet, { zoom: 0.1 });
		expect([g.defaultColWidth, g.defaultRowHeight]).toEqual([6, 2]);
		expect(new SheetGeometry(createWorksheet('S', 1), { zoom: 0.01 }).defaultRowHeight).toBe(1);
	});

	it('scales offsets as sums of rounded lines', () => {
		const g = new SheetGeometry(createWorksheet('S', 1), { zoom: 0.9 });
		// 64 px * 0.9 = 57.6 px rounds to 58 px per column, so ten columns are 580 px, not 576.
		expect(g.colWidth(0)).toBe(58);
		expect(g.rowHeight(0)).toBe(18);
		expect(g.colLeft(10)).toBe(580);
		expect(g.rowTop(10)).toBe(180);
	});
});

describe('SheetGeometry: rectangles and hit testing', () => {
	const sheet = edit((tx, wb) => {
		setColumnsHidden(tx, wb, 0, 0, 0, true);
		setColumnsHidden(tx, wb, 0, 2, 2, true);
		setRowsHidden(tx, wb, 0, 1, 1, true);
	});
	const g = new SheetGeometry(sheet, { zoom: 1 });

	it('measures cell and range rectangles, hidden lines included as zero', () => {
		expect(g.cellRect(2, 3)).toEqual({ x: 64, y: 20, width: 64, height: 20 });
		expect(g.cellRect(1, 2)).toEqual({ x: 64, y: 20, width: 0, height: 0 });
		expect(g.rangeRect({ top: 0, left: 0, bottom: 2, right: 3 })).toEqual({
			x: 0,
			y: 0,
			width: 128,
			height: 40,
		});
	});

	it('finds the cell under a point, skipping hidden lines', () => {
		expect(g.cellAt(0, 0)).toEqual({ row: 0, col: 1 });
		expect(g.cellAt(-10, -10)).toEqual({ row: 0, col: 1 });
		expect(g.cellAt(63, 19)).toEqual({ row: 0, col: 1 });
		// A boundary belongs to the next visible cell.
		expect(g.cellAt(64, 20)).toEqual({ row: 2, col: 3 });
		expect(g.cellAt(127.5, 39.5)).toEqual({ row: 2, col: 3 });
		expect(g.cellAt(1e12, 1e12)).toEqual({ row: 1048575, col: MAX_COL_INDEX });
	});
});

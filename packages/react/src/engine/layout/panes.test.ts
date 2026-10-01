import { describe, expect, it } from 'vitest';
import type { CellAddress, Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	createWorkbook,
	EditTransaction,
	setColumnWidth,
	setFreezePanes,
	setRowHeight,
	setRowsHidden,
} from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import { columnHeaderHeight, computePanes, rowHeaderWidth } from './panes';

function sheetWith(build: (tx: EditTransaction, wb: Workbook) => void): Worksheet {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	build(tx, wb);
	tx.commit();
	return wb.sheets[0];
}

const frozen = (at: CellAddress, origin?: CellAddress) =>
	sheetWith((tx, wb) => setFreezePanes(tx, wb, 0, at, origin));

const panesOf = (sheet: Worksheet, zoom = 1) =>
	computePanes(sheet, new SheetGeometry(sheet, { zoom }));

const NONE = {
	originRow: 0,
	originCol: 0,
	frozenRows: 0,
	frozenCols: 0,
	bodyRow: 0,
	bodyCol: 0,
	frozenHeight: 0,
	frozenWidth: 0,
};

describe('computePanes', () => {
	it('has no frozen band without frozen panes', () => {
		expect(panesOf(sheetWith(() => undefined))).toEqual(NONE);
		const split = sheetWith(() => undefined);
		split.view = { ...split.view, pane: { xSplit: 2000, ySplit: 1500, state: 'split' } };
		expect(panesOf(split)).toEqual(NONE);
	});

	it('freezes rows above and columns left of the frozen cell', () => {
		expect(panesOf(frozen({ row: 2, col: 1 }))).toEqual({
			originRow: 0,
			originCol: 0,
			frozenRows: 2,
			frozenCols: 1,
			bodyRow: 2,
			bodyCol: 1,
			frozenHeight: 40,
			frozenWidth: 64,
		});
	});

	it('freezes only rows or only columns', () => {
		expect(panesOf(frozen({ row: 3, col: 0 }))).toMatchObject({
			frozenRows: 3,
			frozenCols: 0,
			frozenHeight: 60,
			frozenWidth: 0,
			bodyCol: 0,
		});
		expect(panesOf(frozen({ row: 0, col: 2 }))).toMatchObject({
			frozenRows: 0,
			frozenCols: 2,
			frozenHeight: 0,
			frozenWidth: 128,
		});
	});

	it('starts at the top-left cell that was visible when the panes were frozen', () => {
		expect(panesOf(frozen({ row: 7, col: 3 }, { row: 5, col: 1 }))).toEqual({
			originRow: 5,
			originCol: 1,
			frozenRows: 2,
			frozenCols: 2,
			bodyRow: 7,
			bodyCol: 3,
			frozenHeight: 40,
			frozenWidth: 128,
		});
	});

	it('measures the band with custom, hidden and zoomed lines', () => {
		const sheet = sheetWith((tx, wb) => {
			setRowHeight(tx, wb, 0, 0, 0, 30);
			setRowsHidden(tx, wb, 0, 1, 1, true);
			setColumnWidth(tx, wb, 0, 0, 0, 20);
			setFreezePanes(tx, wb, 0, { row: 3, col: 2 });
		});
		expect(panesOf(sheet)).toMatchObject({ frozenHeight: 40 + 0 + 20, frozenWidth: 140 + 64 });
		expect(panesOf(sheet, 2)).toMatchObject({ frozenHeight: 80 + 0 + 40, frozenWidth: 280 + 128 });
	});

	it('treats frozenSplit as frozen and rounds the split counts', () => {
		const sheet = sheetWith(() => undefined);
		sheet.view = { ...sheet.view, pane: { xSplit: 1.6, ySplit: -2, state: 'frozenSplit' } };
		expect(panesOf(sheet)).toMatchObject({
			frozenCols: 2,
			frozenRows: 0,
			bodyCol: 2,
			bodyRow: 0,
			frozenWidth: 128,
		});
	});
});

describe('header sizes', () => {
	it('sizes the row header for at least three digits of the largest row number', () => {
		expect(rowHeaderWidth(1, 1)).toBe(33);
		expect(rowHeaderWidth(100, 1)).toBe(33);
		expect(rowHeaderWidth(1000, 1)).toBe(40);
		expect(rowHeaderWidth(1048576, 1)).toBe(61);
	});

	it('scales header sizes with zoom, but not below 60%', () => {
		expect(rowHeaderWidth(100, 2)).toBe(66);
		expect(rowHeaderWidth(100, 0.6)).toBe(20);
		expect(rowHeaderWidth(100, 0.25)).toBe(20);
		expect(columnHeaderHeight(1)).toBe(20);
		expect(columnHeaderHeight(1.5)).toBe(30);
		expect(columnHeaderHeight(0.5)).toBe(12);
		expect(columnHeaderHeight(0.1)).toBe(12);
	});
});

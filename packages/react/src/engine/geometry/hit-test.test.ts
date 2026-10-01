import { describe, expect, it } from 'vitest';
import type { CellAddress, Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	createWorkbook,
	EditTransaction,
	setColumnsHidden,
	setFreezePanes,
	setRowsHidden,
} from 'xlsx-viewer-core';

import { computePanes } from '../layout/panes';
import {
	hitTest,
	sheetToViewportX,
	sheetToViewportY,
	viewportToSheetX,
	viewportToSheetY,
	type GridMetrics,
} from './hit-test';
import { SheetGeometry } from './sheet-geometry';

const HW = 40;
const HH = 20;

function sheetWith(build?: (tx: EditTransaction, wb: Workbook) => void): Worksheet {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	build?.(tx, wb);
	tx.commit();
	return wb.sheets[0];
}

function frozenAt(at: CellAddress, origin?: CellAddress): Worksheet {
	return sheetWith((tx, wb) => setFreezePanes(tx, wb, 0, at, origin));
}

function metrics(sheet: Worksheet, scrollLeft = 0, scrollTop = 0): GridMetrics {
	const geometry = new SheetGeometry(sheet, { zoom: 1 });
	return {
		geometry,
		panes: computePanes(sheet, geometry),
		headerWidth: HW,
		headerHeight: HH,
		scrollLeft,
		scrollTop,
	};
}

describe('hitTest: headers', () => {
	const m = metrics(sheetWith());

	it('finds the select-all corner', () => {
		expect(hitTest(m, 10, 10)).toMatchObject({ kind: 'corner', row: 0, col: 0 });
		expect(hitTest(m, -5, -5).kind).toBe('corner');
	});

	it('finds column headers by the sheet x under the pointer', () => {
		const hit = hitTest(m, HW + 100, 5);
		expect(hit).toMatchObject({ kind: 'colHeader', col: 1, row: 0, x: 100 });
		expect(hit.colResize).toBeUndefined();
		expect(hit.rowResize).toBeUndefined();
	});

	it('finds row headers by the sheet y under the pointer', () => {
		expect(hitTest(m, 10, HH + 45)).toMatchObject({ kind: 'rowHeader', row: 2, col: 0, y: 45 });
	});

	it('offers to resize the column whose right edge is within 4 px', () => {
		expect(hitTest(m, HW + 126, 5).colResize).toBe(1);
		expect(hitTest(m, HW + 60, 5).colResize).toBe(0);
		expect(hitTest(m, HW + 59, 5).colResize).toBeUndefined();
		// Just right of a boundary resizes the column on its left.
		expect(hitTest(m, HW + 66, 5)).toMatchObject({ col: 1, colResize: 0 });
		expect(hitTest(m, HW + 68, 5).colResize).toBe(0);
		expect(hitTest(m, HW + 69, 5).colResize).toBeUndefined();
		// The left edge of column A is not a resize handle.
		expect(hitTest(m, HW + 2, 5).colResize).toBeUndefined();
	});

	it('offers to resize the row whose bottom edge is within 4 px', () => {
		expect(hitTest(m, 10, HH + 58).rowResize).toBe(2);
		expect(hitTest(m, 10, HH + 41)).toMatchObject({ row: 2, rowResize: 1 });
		expect(hitTest(m, 10, HH + 50).rowResize).toBeUndefined();
		expect(hitTest(m, 10, HH + 1).rowResize).toBeUndefined();
	});

	it('resizes the last visible line before hidden ones', () => {
		const hidden = metrics(
			sheetWith((tx, wb) => {
				setColumnsHidden(tx, wb, 0, 1, 1, true);
				setRowsHidden(tx, wb, 0, 1, 2, true);
			}),
		);
		// Columns: A = 0-64, B hidden, C = 64-128. Rows: 1 = 0-20, 2-3 hidden, 4 = 20-40.
		expect(hitTest(hidden, HW + 66, 5)).toMatchObject({ col: 2, colResize: 0 });
		expect(hitTest(hidden, 10, HH + 22)).toMatchObject({ row: 3, rowResize: 0 });
	});
});

describe('hitTest: cells', () => {
	it('reports the cell and its sheet coordinates in the body', () => {
		const hit = hitTest(metrics(sheetWith()), HW + 130, HH + 45);
		expect(hit).toEqual({
			kind: 'cell',
			row: 2,
			col: 2,
			x: 130,
			y: 45,
			inBodyX: true,
			inBodyY: true,
		});
	});

	it('adds the scroll offset in the scrolled body and in the headers', () => {
		const m = metrics(sheetWith(), 640, 200);
		expect(hitTest(m, HW + 10, HH + 5)).toMatchObject({
			kind: 'cell',
			row: 10,
			col: 10,
			x: 650,
			y: 205,
		});
		expect(hitTest(m, HW + 10, 5)).toMatchObject({ kind: 'colHeader', col: 10 });
		expect(hitTest(m, 5, HH + 5)).toMatchObject({ kind: 'rowHeader', row: 10 });
	});

	it('skips hidden lines', () => {
		const m = metrics(sheetWith((tx, wb) => setColumnsHidden(tx, wb, 0, 0, 0, true)));
		expect(hitTest(m, HW + 1, HH + 1)).toMatchObject({ row: 0, col: 1 });
	});
});

describe('hitTest: frozen panes', () => {
	// Rows 1-2 and columns A-B frozen, body scrolled by 10 columns and 10 rows.
	const m = metrics(frozenAt({ row: 2, col: 2 }), 640, 200);

	it('maps the frozen band without the scroll offset', () => {
		expect(m.panes).toMatchObject({ frozenWidth: 128, frozenHeight: 40, bodyRow: 2, bodyCol: 2 });
		expect(hitTest(m, HW + 100, HH + 30)).toMatchObject({
			kind: 'cell',
			row: 1,
			col: 1,
			inBodyX: false,
			inBodyY: false,
		});
	});

	it('maps the body from the first unfrozen line plus the scroll offset', () => {
		expect(hitTest(m, HW + 130, HH + 45)).toMatchObject({
			row: 12,
			col: 12,
			x: 128 + 640 + 2,
			y: 40 + 200 + 5,
			inBodyX: true,
			inBodyY: true,
		});
		// Frozen rows, scrolled columns.
		expect(hitTest(m, HW + 130, HH + 30)).toMatchObject({
			row: 1,
			col: 12,
			inBodyX: true,
			inBodyY: false,
		});
	});

	it('maps headers over the frozen band', () => {
		expect(hitTest(m, HW + 100, 5)).toMatchObject({ kind: 'colHeader', col: 1 });
		expect(hitTest(m, 5, HH + 30)).toMatchObject({ kind: 'rowHeader', row: 1 });
		expect(hitTest(m, HW + 126, 5).colResize).toBe(1);
	});

	it('starts the frozen band at the top-left cell the panes were frozen from', () => {
		const scrolled = metrics(frozenAt({ row: 7, col: 3 }, { row: 5, col: 1 }));
		expect(scrolled.panes).toMatchObject({
			originRow: 5,
			originCol: 1,
			frozenHeight: 40,
			frozenWidth: 128,
		});
		expect(hitTest(scrolled, HW + 1, HH + 1)).toMatchObject({ row: 5, col: 1 });
		expect(hitTest(scrolled, HW + 130, HH + 45)).toMatchObject({ row: 7, col: 3 });
	});
});

describe('viewport and sheet coordinates', () => {
	const m = metrics(frozenAt({ row: 2, col: 2 }), 640, 200);

	it('converts viewport points to sheet points, flagging the scrollable body', () => {
		expect(viewportToSheetX(m, HW + 100)).toEqual({ x: 100, body: false });
		expect(viewportToSheetX(m, HW + 128)).toEqual({ x: 768, body: true });
		expect(viewportToSheetY(m, HH + 39)).toEqual({ y: 39, body: false });
		expect(viewportToSheetY(m, HH + 40)).toEqual({ y: 240, body: true });
		// Without frozen panes everything right of the header is body.
		const plain = metrics(sheetWith(), 10, 0);
		expect(viewportToSheetX(plain, HW)).toEqual({ x: 10, body: true });
		expect(viewportToSheetX(plain, 0)).toEqual({ x: -40, body: false });
	});

	it('converts sheet points back to viewport points for frozen and body lines', () => {
		for (const px of [HW + 5, HW + 100, HW + 128, HW + 300]) {
			const { x } = viewportToSheetX(m, px);
			expect(sheetToViewportX(m, x, m.geometry.cols.indexAt(x))).toBe(px);
		}
		for (const py of [HH + 5, HH + 39, HH + 40, HH + 400]) {
			const { y } = viewportToSheetY(m, py);
			expect(sheetToViewportY(m, y, m.geometry.rows.indexAt(y))).toBe(py);
		}
		// Body cells scrolled out of view land left of / above the frozen band.
		expect(sheetToViewportX(m, m.geometry.colLeft(3), 3)).toBe(HW + 128 + 64 - 640);
		expect(sheetToViewportY(m, m.geometry.rowTop(2), 2)).toBe(HH + 40 - 200);
	});
});

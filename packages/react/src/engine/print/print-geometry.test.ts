import { describe, expect, it } from 'vitest';
import type { Worksheet } from 'xlsx-viewer-core';
import {
	cellKey,
	createWorkbook,
	createWorksheet,
	MAX_COL_INDEX,
	printColWidth,
	printRowHeight,
	StyleResolver,
} from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import { MergeIndex } from '../layout/style-at';
import { layoutViewport } from '../layout/viewport-layout';
import { createPrintGeometry, PX_PER_PT } from './print-geometry';

function sheet(): Worksheet {
	const s = createWorksheet('S', 1);
	s.rows.set(1, { height: 14.4, customHeight: true });
	s.rows.set(2, { hidden: true });
	s.rows.set(3, { customFormat: true, style: 0 });
	s.cols = [
		{ min: 1, max: 3, width: 20 },
		{ min: 4, max: 4, hidden: true },
	];
	return s;
}

describe('createPrintGeometry', () => {
	it('uses the printed sizes in points at 96 DPI', () => {
		const g = createPrintGeometry(sheet(), 1);
		expect(PX_PER_PT).toBeCloseTo(4 / 3, 12);
		expect(g.zoom).toBe(1);
		expect(g.defaultRowHeight).toBe(20);
		expect(g.defaultColWidth).toBe(64);
		expect([0, 2, 3, 4].map((r) => g.rowHeight(r))).toEqual([20, 0, 20, 20]);
		expect(g.rowHeight(1)).toBeCloseTo(19.2, 9);
		expect([0, 1, 2, 3, 4, 5].map((c) => g.colWidth(c))).toEqual([64, 140, 140, 140, 0, 64]);
	});

	it('matches the core print metrics at any scale, without per-line rounding', () => {
		const s = sheet();
		for (const scale of [0.37, 0.5, 1, 1.85]) {
			const g = createPrintGeometry(s, scale);
			for (let row = 0; row < 5; row++) {
				expect(g.rowHeight(row)).toBeCloseTo(printRowHeight(s, row) * PX_PER_PT * scale, 9);
			}
			for (let col = 0; col < 6; col++) {
				expect(g.colWidth(col)).toBeCloseTo(printColWidth(s, col) * PX_PER_PT * scale, 9);
			}
		}
		const small = createPrintGeometry(s, 0.37);
		expect(small.colWidth(1)).toBeCloseTo(51.8, 9);
		expect(small.colLeft(4)).toBeCloseTo((48 + 3 * 105) * PX_PER_PT * 0.37, 9);
		// The screen grid rounds every line instead.
		expect(new SheetGeometry(s, { zoom: 0.37 }).colWidth(1)).toBe(52);
	});

	it('accumulates offsets over hidden and custom lines', () => {
		const g = createPrintGeometry(sheet(), 0.5);
		expect(g.rowTop(4)).toBeCloseTo(10 + 9.6 + 0 + 10, 9);
		expect(g.colLeft(6)).toBe(32 + 3 * 70 + 0 + 32);
	});

	it('hides rows without a record when rows are hidden by default', () => {
		const s = sheet();
		s.zeroHeight = true;
		const g = createPrintGeometry(s, 1);
		expect(g.defaultRowHeight).toBe(0);
		expect([0, 2, 3, 4].map((r) => g.rowHeight(r))).toEqual([0, 0, 20, 0]);
		expect(g.rowHeight(1)).toBeCloseTo(19.2, 9);
	});

	it('clamps column records to the last column', () => {
		const s = createWorksheet('S', 1);
		s.cols = [{ min: 16380, max: 20000, width: 10.7109375 }];
		const g = createPrintGeometry(s, 1);
		expect(g.colWidth(MAX_COL_INDEX)).toBe(75);
		expect(g.colWidth(16379)).toBe(64);
	});

	it('provides the SheetGeometry rectangle and hit-test helpers', () => {
		const g = createPrintGeometry(sheet(), 0.5);
		const rect = g.rangeRect({ top: 0, left: 0, bottom: 1, right: 1 });
		expect(rect).toMatchObject({ x: 0, y: 0, width: 32 + 70 });
		expect(rect.height).toBeCloseTo(10 + 9.6, 9);
		expect(g.cellRect(1, 1)).toEqual(g.rangeRect({ top: 1, left: 1, bottom: 1, right: 1 }));
		expect(g.cellAt(40, 12)).toEqual({ row: 1, col: 1 });
		// Hidden row 3 takes no space: y 25 is inside row 4.
		expect(g.cellAt(0, 25)).toEqual({ row: 3, col: 0 });
		expect(g.rows.isHidden(2)).toBe(true);
	});

	it('lays out cells at their printed positions', () => {
		const wb = createWorkbook();
		const s = wb.sheets[0];
		s.cols = [{ min: 0, max: 0, width: 20 }];
		s.cells.set(cellKey(1, 1), { value: 'B2' });
		const g = createPrintGeometry(s, 0.5);
		const layout = layoutViewport({
			sheet: s,
			geometry: g,
			rows: { start: 0, end: 3 },
			cols: { start: 0, end: 3 },
			styles: new StyleResolver(wb.styles, wb.theme),
			merges: new MergeIndex(s.merges),
			describe: {
				zoom: 0.5,
				colorCtx: { theme: wb.theme },
				date1904: false,
				defaultTextColor: '#000',
			},
		});
		expect(layout.texts.map((t) => [t.key, t.cell])).toEqual([
			['t:1:1', { x: 70, y: 10, width: 32, height: 10 }],
		]);
	});
});

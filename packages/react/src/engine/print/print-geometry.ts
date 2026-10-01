import type { Worksheet } from 'xlsx-viewer-core';
import {
	defaultPrintColWidth,
	defaultPrintRowHeight,
	MAX_COLS,
	MAX_ROWS,
	printColWidth,
	printRowHeight,
} from 'xlsx-viewer-core';

import { AxisMetrics } from '../geometry/axis';
import type { Rect, SheetGeometry } from '../geometry/sheet-geometry';

/** CSS pixels per point (96 DPI). */
export const PX_PER_PT = 96 / 72;

/**
 * Sheet geometry for printing: the core's printed row heights and column
 * widths (points) at a print scale, in CSS pixels at 96 DPI, without the
 * per-line rounding of the screen grid. Cells therefore land exactly where
 * the paginator put them, whatever the scale. Same shape as the grid's
 * `SheetGeometry`, so the grid's layers and drawings can render with it.
 */
export function createPrintGeometry(sheet: Worksheet, scale: number): SheetGeometry {
	const k = PX_PER_PT * scale;
	const defaultRowHeight = defaultPrintRowHeight(sheet) * k;
	const rowSizes = new Map<number, number>();
	for (const row of sheet.rows.keys()) {
		rowSizes.set(row, printRowHeight(sheet, row) * k);
	}
	const rows = new AxisMetrics(defaultRowHeight, rowSizes, MAX_ROWS);
	const defaultColWidth = defaultPrintColWidth(sheet) * k;
	const colSizes = new Map<number, number>();
	for (const c of sheet.cols) {
		const hi = Math.min(c.max, MAX_COLS - 1);
		for (let col = c.min; col <= hi; col++) {
			colSizes.set(col, printColWidth(sheet, col) * k);
		}
	}
	const cols = new AxisMetrics(defaultColWidth, colSizes, MAX_COLS);
	const rangeRect = (r: { top: number; left: number; bottom: number; right: number }): Rect => {
		const x = cols.offset(r.left);
		const y = rows.offset(r.top);
		return { x, y, width: cols.end(r.right) - x, height: rows.end(r.bottom) - y };
	};
	return {
		rows,
		cols,
		zoom: scale,
		defaultRowHeight,
		defaultColWidth,
		colLeft: (col) => cols.offset(col),
		colWidth: (col) => cols.size(col),
		rowTop: (row) => rows.offset(row),
		rowHeight: (row) => rows.size(row),
		rangeRect,
		cellRect: (row, col) => rangeRect({ top: row, left: col, bottom: row, right: col }),
		cellAt: (x, y) => ({ row: rows.indexAt(y), col: cols.indexAt(x) }),
	};
}

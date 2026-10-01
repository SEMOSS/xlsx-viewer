import type { RangeAddress, Worksheet } from 'xlsx-viewer-core';
import {
	columnWidthToPixels,
	defaultColumnWidthPx,
	MAX_COLS,
	MAX_ROWS,
	pointsToPixels,
} from 'xlsx-viewer-core';

import { AxisMetrics } from './axis';

export interface Rect {
	x: number;
	y: number;
	width: number;
	height: number;
}

export interface GeometryOptions {
	/** Zoom factor (1 = 100%). */
	zoom: number;
	/** Maximum digit width of the workbook's default font in px at 100%. */
	maxDigitWidth?: number;
}

/**
 * Pixel geometry of a worksheet at a zoom level: row / column sizes and
 * offsets (in zoomed CSS pixels, origin at A1's top-left corner, excluding
 * headers), hit testing and range rectangles.
 */
export class SheetGeometry {
	readonly rows: AxisMetrics;
	readonly cols: AxisMetrics;
	readonly zoom: number;
	/** Default row height at this zoom. */
	readonly defaultRowHeight: number;
	readonly defaultColWidth: number;

	constructor(sheet: Worksheet, options: GeometryOptions) {
		const zoom = options.zoom;
		const mdw = options.maxDigitWidth ?? 7;
		this.zoom = zoom;

		const defRowPx = Math.round(pointsToPixels(sheet.defaultRowHeight || 15));
		this.defaultRowHeight = sheet.zeroHeight ? 0 : Math.max(1, Math.round(defRowPx * zoom));
		const rowOverrides = new Map<number, number>();
		for (const [r, info] of sheet.rows) {
			if (info.hidden) {
				rowOverrides.set(r, 0);
			} else if (info.height !== undefined) {
				rowOverrides.set(
					r,
					Math.max(0, Math.round(Math.round(pointsToPixels(info.height)) * zoom)),
				);
			} else if (sheet.zeroHeight) {
				rowOverrides.set(r, Math.max(1, Math.round(defRowPx * zoom)));
			}
		}
		this.rows = new AxisMetrics(this.defaultRowHeight, rowOverrides, MAX_ROWS);

		const defColPx = defaultColumnWidthPx(sheet.defaultColWidth, sheet.baseColWidth, mdw);
		this.defaultColWidth = Math.max(1, Math.round(defColPx * zoom));
		const colOverrides = new Map<number, number>();
		for (const c of sheet.cols) {
			let px: number | undefined;
			if (c.hidden) {
				px = 0;
			} else if (c.width !== undefined) {
				px = Math.round(columnWidthToPixels(c.width, mdw) * zoom);
			}
			if (px === undefined) {
				continue;
			}
			const hi = Math.min(c.max, MAX_COLS - 1);
			for (let col = c.min; col <= hi; col++) {
				colOverrides.set(col, px);
			}
		}
		this.cols = new AxisMetrics(this.defaultColWidth, colOverrides, MAX_COLS);
	}

	colLeft(col: number): number {
		return this.cols.offset(col);
	}

	colWidth(col: number): number {
		return this.cols.size(col);
	}

	rowTop(row: number): number {
		return this.rows.offset(row);
	}

	rowHeight(row: number): number {
		return this.rows.size(row);
	}

	/** Rectangle of a range (merged areas included as given). */
	rangeRect(r: RangeAddress): Rect {
		const x = this.cols.offset(r.left);
		const y = this.rows.offset(r.top);
		return {
			x,
			y,
			width: this.cols.end(r.right) - x,
			height: this.rows.end(r.bottom) - y,
		};
	}

	cellRect(row: number, col: number): Rect {
		return {
			x: this.cols.offset(col),
			y: this.rows.offset(row),
			width: this.cols.size(col),
			height: this.rows.size(row),
		};
	}

	/** Cell under a point in sheet coordinates. */
	cellAt(x: number, y: number): { row: number; col: number } {
		return { row: this.rows.indexAt(y), col: this.cols.indexAt(x) };
	}
}

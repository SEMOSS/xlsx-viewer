import type { PaneLayout } from '../layout/panes';
import type { SheetGeometry } from './sheet-geometry';

export interface GridMetrics {
	geometry: SheetGeometry;
	panes: PaneLayout;
	headerWidth: number;
	headerHeight: number;
	scrollLeft: number;
	scrollTop: number;
}

export type HitKind = 'corner' | 'colHeader' | 'rowHeader' | 'cell';

export interface GridHit {
	kind: HitKind;
	row: number;
	col: number;
	/** Point in sheet coordinates. */
	x: number;
	y: number;
	/** Column whose right edge is under the pointer (header resize). */
	colResize?: number;
	/** Row whose bottom edge is under the pointer (header resize). */
	rowResize?: number;
	/** The point is in the scrollable body region (not frozen / headers). */
	inBodyX: boolean;
	inBodyY: boolean;
}

const RESIZE_SLOP = 4;

/** Convert a viewport x (relative to the scroller's client area) to a sheet x. */
export function viewportToSheetX(m: GridMetrics, px: number): { x: number; body: boolean } {
	const { geometry: g, panes } = m;
	if (px < m.headerWidth + panes.frozenWidth) {
		return { x: g.colLeft(panes.originCol) + (px - m.headerWidth), body: false };
	}
	return {
		x: g.colLeft(panes.bodyCol) + m.scrollLeft + (px - m.headerWidth - panes.frozenWidth),
		body: true,
	};
}

export function viewportToSheetY(m: GridMetrics, py: number): { y: number; body: boolean } {
	const { geometry: g, panes } = m;
	if (py < m.headerHeight + panes.frozenHeight) {
		return { y: g.rowTop(panes.originRow) + (py - m.headerHeight), body: false };
	}
	return {
		y: g.rowTop(panes.bodyRow) + m.scrollTop + (py - m.headerHeight - panes.frozenHeight),
		body: true,
	};
}

/** Sheet x to viewport x (inverse of {@link viewportToSheetX}), for frozen or body columns. */
export function sheetToViewportX(m: GridMetrics, x: number, col: number): number {
	const { geometry: g, panes } = m;
	if (col < panes.bodyCol) {
		return m.headerWidth + (x - g.colLeft(panes.originCol));
	}
	return m.headerWidth + panes.frozenWidth + (x - g.colLeft(panes.bodyCol) - m.scrollLeft);
}

export function sheetToViewportY(m: GridMetrics, y: number, row: number): number {
	const { geometry: g, panes } = m;
	if (row < panes.bodyRow) {
		return m.headerHeight + (y - g.rowTop(panes.originRow));
	}
	return m.headerHeight + panes.frozenHeight + (y - g.rowTop(panes.bodyRow) - m.scrollTop);
}

/** What lies under a viewport point. */
export function hitTest(m: GridMetrics, px: number, py: number): GridHit {
	const { geometry: g } = m;
	const sx = viewportToSheetX(m, Math.max(px, m.headerWidth));
	const sy = viewportToSheetY(m, Math.max(py, m.headerHeight));
	const col = g.cols.indexAt(sx.x);
	const row = g.rows.indexAt(sy.y);
	const inHeaderX = px < m.headerWidth;
	const inHeaderY = py < m.headerHeight;
	const base = { row, col, x: sx.x, y: sy.y, inBodyX: sx.body, inBodyY: sy.body };
	if (inHeaderX && inHeaderY) {
		return { kind: 'corner', ...base };
	}
	if (inHeaderY) {
		const hit: GridHit = { kind: 'colHeader', ...base };
		const right = g.colLeft(col) + g.colWidth(col);
		if (right - sx.x <= RESIZE_SLOP) {
			hit.colResize = col;
		} else if (sx.x - g.colLeft(col) <= RESIZE_SLOP && col > 0) {
			hit.colResize = g.cols.prevVisible(col - 1);
		}
		return hit;
	}
	if (inHeaderX) {
		const hit: GridHit = { kind: 'rowHeader', ...base };
		const bottom = g.rowTop(row) + g.rowHeight(row);
		if (bottom - sy.y <= RESIZE_SLOP) {
			hit.rowResize = row;
		} else if (sy.y - g.rowTop(row) <= RESIZE_SLOP && row > 0) {
			hit.rowResize = g.rows.prevVisible(row - 1);
		}
		return hit;
	}
	return { kind: 'cell', ...base };
}

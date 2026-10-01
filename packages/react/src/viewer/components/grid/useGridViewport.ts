import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

import { MAX_COLS, MAX_ROWS } from 'xlsx-viewer-core';

import type { GridMetrics } from '../../../engine/geometry/hit-test';
import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { PaneLayout } from '../../../engine/layout/panes';
import type { Span } from '../../../engine/layout/types';

const OVERSCAN_ROWS = 12;
const OVERSCAN_COLS = 4;

export interface ViewportState {
	width: number;
	height: number;
	rows: Span;
	cols: Span;
	/** Scrollable extent in rows / columns. */
	extentRows: number;
	extentCols: number;
}

interface Options {
	geometry: SheetGeometry;
	panes: PaneLayout;
	headerWidth: number;
	headerHeight: number;
	/** Used range bottom-right, to size the initial scroll extent. */
	usedRows: number;
	usedCols: number;
	sheetId: string;
	initialScroll?: { left: number; top: number };
	onScrollPosition?: (pos: { left: number; top: number }) => void;
}

function visibleSpan(
	axis: SheetGeometry['rows'],
	start: number,
	origin: number,
	scroll: number,
	size: number,
	overscan: number,
	limit: number,
): Span {
	const first = axis.indexAt(axis.offset(origin) + scroll);
	const last = axis.indexAt(axis.offset(origin) + scroll + Math.max(0, size));
	return {
		start: Math.max(start, first - overscan),
		end: Math.min(limit - 1, last + overscan),
	};
}

/**
 * Scroll state of the grid: the visible (overscanned) body window, the
 * scrollable extent and helpers to read live metrics or scroll a cell into
 * view. Re-renders only when the window must move.
 */
export function useGridViewport(opts: Options) {
	const scrollerRef = useRef<HTMLDivElement | null>(null);
	const scrollPos = useRef({ left: 0, top: 0 });
	const optsRef = useRef(opts);
	optsRef.current = opts;
	const [state, setState] = useState<ViewportState>(() => ({
		width: 800,
		height: 600,
		rows: { start: opts.panes.bodyRow, end: opts.panes.bodyRow + 40 },
		cols: { start: opts.panes.bodyCol, end: opts.panes.bodyCol + 15 },
		extentRows: Math.min(MAX_ROWS, Math.max(opts.usedRows + 100, 200)),
		extentCols: Math.min(MAX_COLS, Math.max(opts.usedCols + 20, 40)),
	}));
	const stateRef = useRef(state);
	stateRef.current = state;

	const recompute = useCallback((force = false) => {
		const o = optsRef.current;
		const el = scrollerRef.current;
		if (!el) {
			return;
		}
		const width = el.clientWidth;
		const height = el.clientHeight;
		const { left, top } = scrollPos.current;
		const bodyW = width - o.headerWidth - o.panes.frozenWidth;
		const bodyH = height - o.headerHeight - o.panes.frozenHeight;
		const cur = stateRef.current;
		// Strictly visible spans; the rendered window keeps an overscan margin around them.
		const vis = visibleSpan(
			o.geometry.rows,
			o.panes.bodyRow,
			o.panes.bodyRow,
			top,
			bodyH,
			0,
			MAX_ROWS,
		);
		const visC = visibleSpan(
			o.geometry.cols,
			o.panes.bodyCol,
			o.panes.bodyCol,
			left,
			bodyW,
			0,
			MAX_COLS,
		);
		let extentRows = cur.extentRows;
		let extentCols = cur.extentCols;
		const visRows = vis.end - vis.start + 1;
		const visCols = visC.end - visC.start + 1;
		if (vis.end + visRows > extentRows) {
			extentRows = Math.min(MAX_ROWS, Math.max(extentRows + visRows * 3, vis.end + visRows * 2));
		}
		if (visC.end + visCols > extentCols) {
			extentCols = Math.min(MAX_COLS, Math.max(extentCols + visCols * 2, visC.end + visCols * 2));
		}
		const covered =
			cur.rows.start <= vis.start &&
			cur.rows.end >= vis.end &&
			cur.cols.start <= visC.start &&
			cur.cols.end >= visC.end;
		if (
			!force &&
			covered &&
			width === cur.width &&
			height === cur.height &&
			extentRows === cur.extentRows &&
			extentCols === cur.extentCols
		) {
			return;
		}
		const rows = {
			start: Math.max(o.panes.bodyRow, vis.start - OVERSCAN_ROWS),
			end: Math.min(MAX_ROWS - 1, vis.end + OVERSCAN_ROWS),
		};
		const cols = {
			start: Math.max(o.panes.bodyCol, visC.start - OVERSCAN_COLS),
			end: Math.min(MAX_COLS - 1, visC.end + OVERSCAN_COLS),
		};
		setState({ width, height, rows, cols, extentRows, extentCols });
	}, []);

	const onScroll = useCallback(() => {
		const el = scrollerRef.current;
		if (!el) {
			return;
		}
		scrollPos.current = { left: el.scrollLeft, top: el.scrollTop };
		optsRef.current.onScrollPosition?.(scrollPos.current);
		recompute();
	}, [recompute]);

	useLayoutEffect(() => {
		const el = scrollerRef.current;
		if (!el) {
			return;
		}
		const ro = new ResizeObserver(() => recompute());
		ro.observe(el);
		recompute(true);
		return () => ro.disconnect();
	}, [recompute]);

	// Restore the sheet's scroll position when switching sheets.
	useLayoutEffect(() => {
		const el = scrollerRef.current;
		if (!el) {
			return;
		}
		const pos = opts.initialScroll ?? { left: 0, top: 0 };
		el.scrollLeft = pos.left;
		el.scrollTop = pos.top;
		scrollPos.current = { left: el.scrollLeft, top: el.scrollTop };
		recompute(true);
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [opts.sheetId]);

	// Geometry / frozen panes changed (zoom, row heights...): recompute the window.
	useEffect(() => {
		recompute(true);
	}, [opts.geometry, opts.panes, opts.headerWidth, opts.headerHeight, recompute]);

	const metrics = useCallback((): GridMetrics => {
		const o = optsRef.current;
		return {
			geometry: o.geometry,
			panes: o.panes,
			headerWidth: o.headerWidth,
			headerHeight: o.headerHeight,
			scrollLeft: scrollPos.current.left,
			scrollTop: scrollPos.current.top,
		};
	}, []);

	/** Scroll so that the cell (or range corner) is visible in the body. */
	/** Scroll a cell into view; a `null` row (column) leaves that axis where it is. */
	const ensureVisible = useCallback((row: number | null, col: number | null) => {
		const el = scrollerRef.current;
		if (!el) {
			return;
		}
		const o = optsRef.current;
		const g = o.geometry;
		const bodyW = el.clientWidth - o.headerWidth - o.panes.frozenWidth;
		const bodyH = el.clientHeight - o.headerHeight - o.panes.frozenHeight;
		let { left, top } = scrollPos.current;
		if (col !== null && col >= o.panes.bodyCol) {
			const x = g.colLeft(col) - g.colLeft(o.panes.bodyCol);
			const w = g.colWidth(col);
			if (x < left) {
				left = x;
			} else if (x + w > left + bodyW) {
				left = Math.max(0, Math.min(x, x + w - bodyW));
			}
		}
		if (row !== null && row >= o.panes.bodyRow) {
			const y = g.rowTop(row) - g.rowTop(o.panes.bodyRow);
			const h = g.rowHeight(row);
			if (y < top) {
				top = y;
			} else if (y + h > top + bodyH) {
				top = Math.max(0, Math.min(y, y + h - bodyH));
			}
		}
		const cur = stateRef.current;
		const r = row ?? 0;
		const c = col ?? 0;
		if (r + 40 > cur.extentRows || c + 10 > cur.extentCols) {
			setState({
				...cur,
				extentRows: Math.min(MAX_ROWS, Math.max(cur.extentRows, r + 200)),
				extentCols: Math.min(MAX_COLS, Math.max(cur.extentCols, c + 30)),
			});
		}
		if (left !== scrollPos.current.left || top !== scrollPos.current.top) {
			// Apply after a possible extent growth has been rendered.
			requestAnimationFrame(() => {
				el.scrollLeft = left;
				el.scrollTop = top;
			});
			el.scrollLeft = left;
			el.scrollTop = top;
		}
	}, []);

	const scrollBy = useCallback((dx: number, dy: number) => {
		const el = scrollerRef.current;
		if (el) {
			el.scrollLeft += dx;
			el.scrollTop += dy;
		}
	}, []);

	return { scrollerRef, state, onScroll, metrics, ensureVisible, scrollBy };
}

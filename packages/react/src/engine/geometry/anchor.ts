import type { AnchorPoint, DrawingAnchor } from 'xlsx-viewer-core';
import { EMU_PER_PIXEL } from 'xlsx-viewer-core';

import type { Rect, SheetGeometry } from './sheet-geometry';

function pointToPx(p: AnchorPoint, g: SheetGeometry): { x: number; y: number } {
	const colOff = Math.min((p.colOff / EMU_PER_PIXEL) * g.zoom, g.colWidth(p.col));
	const rowOff = Math.min((p.rowOff / EMU_PER_PIXEL) * g.zoom, g.rowHeight(p.row));
	return { x: g.colLeft(p.col) + colOff, y: g.rowTop(p.row) + rowOff };
}

/** Pixel rectangle (sheet coordinates, zoomed) of a drawing anchor. */
export function anchorRect(anchor: DrawingAnchor, g: SheetGeometry): Rect {
	switch (anchor.type) {
		case 'twoCell': {
			const a = pointToPx(anchor.from, g);
			const b = pointToPx(anchor.to, g);
			return { x: a.x, y: a.y, width: Math.max(0, b.x - a.x), height: Math.max(0, b.y - a.y) };
		}
		case 'oneCell': {
			const a = pointToPx(anchor.from, g);
			return {
				x: a.x,
				y: a.y,
				width: (anchor.ext.cx / EMU_PER_PIXEL) * g.zoom,
				height: (anchor.ext.cy / EMU_PER_PIXEL) * g.zoom,
			};
		}
		case 'absolute':
			return {
				x: (anchor.pos.x / EMU_PER_PIXEL) * g.zoom,
				y: (anchor.pos.y / EMU_PER_PIXEL) * g.zoom,
				width: (anchor.ext.cx / EMU_PER_PIXEL) * g.zoom,
				height: (anchor.ext.cy / EMU_PER_PIXEL) * g.zoom,
			};
	}
}

function pxToPoint(x: number, y: number, g: SheetGeometry): AnchorPoint {
	const col = g.cols.indexAt(x);
	const row = g.rows.indexAt(y);
	return {
		col,
		colOff: Math.max(0, Math.round(((x - g.colLeft(col)) / g.zoom) * EMU_PER_PIXEL)),
		row,
		rowOff: Math.max(0, Math.round(((y - g.rowTop(row)) / g.zoom) * EMU_PER_PIXEL)),
	};
}

/** Anchor of the same type as `previous` that places the object at `rect` (sheet px, zoomed). */
export function rectToAnchor(rect: Rect, previous: DrawingAnchor, g: SheetGeometry): DrawingAnchor {
	const ext = {
		cx: Math.round((rect.width / g.zoom) * EMU_PER_PIXEL),
		cy: Math.round((rect.height / g.zoom) * EMU_PER_PIXEL),
	};
	switch (previous.type) {
		case 'twoCell':
			return {
				type: 'twoCell',
				editAs: previous.editAs,
				from: pxToPoint(rect.x, rect.y, g),
				to: pxToPoint(rect.x + rect.width, rect.y + rect.height, g),
			};
		case 'oneCell':
			return { type: 'oneCell', from: pxToPoint(rect.x, rect.y, g), ext };
		case 'absolute':
			return {
				type: 'absolute',
				pos: {
					x: Math.round((rect.x / g.zoom) * EMU_PER_PIXEL),
					y: Math.round((rect.y / g.zoom) * EMU_PER_PIXEL),
				},
				ext,
			};
	}
}

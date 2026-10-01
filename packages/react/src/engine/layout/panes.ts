import type { Worksheet } from 'xlsx-viewer-core';

import type { SheetGeometry } from '../geometry/sheet-geometry';

/** Frozen-pane arrangement of a sheet view. */
export interface PaneLayout {
	/** First displayed row / column (rows above a frozen origin are not shown at all). */
	originRow: number;
	originCol: number;
	/** Number of frozen rows / columns (0 when not frozen). */
	frozenRows: number;
	frozenCols: number;
	/** First row / column of the scrollable body. */
	bodyRow: number;
	bodyCol: number;
	/** Pixel size of the frozen bands. */
	frozenHeight: number;
	frozenWidth: number;
}

export function computePanes(sheet: Worksheet, g: SheetGeometry): PaneLayout {
	const pane = sheet.view.pane;
	const frozen = pane && (pane.state === 'frozen' || pane.state === 'frozenSplit');
	if (!frozen) {
		return {
			originRow: 0,
			originCol: 0,
			frozenRows: 0,
			frozenCols: 0,
			bodyRow: 0,
			bodyCol: 0,
			frozenHeight: 0,
			frozenWidth: 0,
		};
	}
	const originRow = sheet.view.topLeftCell?.row ?? 0;
	const originCol = sheet.view.topLeftCell?.col ?? 0;
	const frozenRows = Math.max(0, Math.round(pane.ySplit));
	const frozenCols = Math.max(0, Math.round(pane.xSplit));
	const bodyRow = originRow + frozenRows;
	const bodyCol = originCol + frozenCols;
	return {
		originRow,
		originCol,
		frozenRows,
		frozenCols,
		bodyRow,
		bodyCol,
		frozenHeight: g.rowTop(bodyRow) - g.rowTop(originRow),
		frozenWidth: g.colLeft(bodyCol) - g.colLeft(originCol),
	};
}

/** Width of the row-number header for the largest row number shown. */
export function rowHeaderWidth(maxRowNumber: number, zoom: number): number {
	const digits = Math.max(3, String(maxRowNumber).length);
	return Math.round((digits * 7 + 12) * Math.max(zoom, 0.6));
}

/** Height of the column-letter header. */
export function columnHeaderHeight(zoom: number): number {
	return Math.round(20 * Math.max(zoom, 0.6));
}

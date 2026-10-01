import type { Worksheet } from 'xlsx-viewer-core';
import {
	cellKey,
	keyCol,
	keyRow,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	pixelsToColumnWidth,
	setColumnWidth,
	setColumnsHidden,
	setAutoRowHeight,
	setRowHeight,
	setRowsHidden,
} from 'xlsx-viewer-core';

import { describeCellContent } from '../../engine/render/cell-content';
import { canvasFont } from '../../engine/render/fonts';
import { measureLines } from '../../engine/render/text-measure';
import { CELL_PADDING_X } from '../../engine/layout/viewport-layout';
import { measureRowContent } from '../../engine/layout/row-height';
import { SheetGeometry } from '../../engine/geometry/sheet-geometry';
import type { ViewerStore } from '../state/viewer-store';

/** Columns a header action applies to: the whole-column selection containing `col`, else just `col`. */
function targetColumns(store: ViewerStore, col: number): Array<[number, number]> {
	const ranges = store
		.selectionOf()
		.ranges.filter(
			(r) => r.top === 0 && r.bottom === MAX_ROW_INDEX && col >= r.left && col <= r.right,
		);
	if (ranges.length === 0) {
		return [[col, col]];
	}
	return store
		.selectionOf()
		.ranges.filter((r) => r.top === 0 && r.bottom === MAX_ROW_INDEX)
		.map((r) => [r.left, r.right]);
}

function targetRows(store: ViewerStore, row: number): Array<[number, number]> {
	const ranges = store
		.selectionOf()
		.ranges.filter(
			(r) => r.left === 0 && r.right === MAX_COL_INDEX && row >= r.top && row <= r.bottom,
		);
	if (ranges.length === 0) {
		return [[row, row]];
	}
	return store
		.selectionOf()
		.ranges.filter((r) => r.left === 0 && r.right === MAX_COL_INDEX)
		.map((r) => [r.top, r.bottom]);
}

function zoomOf(sheet: Worksheet): number {
	return (sheet.view.zoomScale || 100) / 100;
}

/** Column resize from a header drag (`widthPx` at the current zoom). */
export function resizeColumns(store: ViewerStore, col: number, widthPx: number): void {
	const sheet = store.activeSheet;
	const px = widthPx / zoomOf(sheet);
	const width = pixelsToColumnWidth(px);
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('undo.columnWidth'), (tx) => {
		for (const [a, b] of targetColumns(store, col)) {
			if (width <= 0) {
				setColumnsHidden(tx, store.workbook, sheetIndex, a, b, true);
			} else {
				setColumnWidth(tx, store.workbook, sheetIndex, a, b, width);
			}
		}
	});
}

/** Row resize from a header drag (`heightPx` at the current zoom). */
export function resizeRows(store: ViewerStore, row: number, heightPx: number): void {
	const sheet = store.activeSheet;
	const pt = Math.round((heightPx / zoomOf(sheet)) * 0.75 * 4) / 4;
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('undo.rowHeight'), (tx) => {
		for (const [a, b] of targetRows(store, row)) {
			if (pt <= 0) {
				setRowsHidden(tx, store.workbook, sheetIndex, a, b, true);
			} else {
				setRowHeight(tx, store.workbook, sheetIndex, a, b, pt);
			}
		}
	});
}

/** Width (px at 100%) needed by the widest content of a column. */
export function measureColumnContent(store: ViewerStore, sheet: Worksheet, col: number): number {
	let max = 0;
	for (const [key, cell] of sheet.cells) {
		if (keyCol(key) !== col || cell.value === null) {
			continue;
		}
		const row = keyRow(key);
		if (
			sheet.merges.some(
				(m) =>
					row >= m.top && row <= m.bottom && col >= m.left && col <= m.right && m.left !== m.right,
			)
		) {
			continue;
		}
		const style = store.styles.resolve(cell.style);
		const content = describeCellContent(cell, style, {
			zoom: 1,
			colorCtx: { theme: store.workbook.theme },
			date1904: store.workbook.date1904,
			defaultTextColor: '#000',
			showFormulas: sheet.view.showFormulas,
		});
		if (!content || content.wrap) {
			continue;
		}
		const w =
			measureLines(content.text, canvasFont(content.font)) +
			content.indent +
			CELL_PADDING_X * 2 +
			3;
		max = Math.max(max, w);
	}
	return Math.ceil(max);
}

/** Double-click on a column boundary: fit the column(s) to their content. */
export function autoFitColumns(store: ViewerStore, col: number): void {
	const sheet = store.activeSheet;
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.autoFitColumn'), (tx) => {
		for (const [a, b] of targetColumns(store, col)) {
			for (let c = a; c <= Math.min(b, a + 200); c++) {
				const px = measureColumnContent(store, sheet, c);
				if (px > 0) {
					setColumnWidth(
						tx,
						store.workbook,
						sheetIndex,
						c,
						c,
						pixelsToColumnWidth(Math.max(px, 12)),
					);
				}
			}
		}
	});
}

/** Double-click on a row boundary: fit the row(s) to their content. */
export function autoFitRows(store: ViewerStore, row: number): void {
	const sheet = store.activeSheet;
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.autoFitRow'), (tx) => {
		const geometry = new SheetGeometry(sheet, { zoom: 1 });
		for (const [a, b] of targetRows(store, row)) {
			for (let r = a; r <= Math.min(b, a + 2000); r++) {
				const pt = measureRowContent(store.workbook, store.styles, sheet, r, geometry);
				setAutoRowHeight(
					tx,
					store.workbook,
					sheetIndex,
					r,
					pt > sheet.defaultRowHeight ? pt : null,
				);
			}
		}
	});
}

export function hideColumns(store: ViewerStore, hidden: boolean): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t(hidden ? 'cmd.hideColumns' : 'cmd.unhideColumns'), (tx) => {
		for (const r of store.selectionOf().ranges) {
			setColumnsHidden(tx, store.workbook, sheetIndex, r.left, r.right, hidden);
		}
	});
}

export function hideRows(store: ViewerStore, hidden: boolean): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t(hidden ? 'cmd.hideRows' : 'cmd.unhideRows'), (tx) => {
		for (const r of store.selectionOf().ranges) {
			setRowsHidden(
				tx,
				store.workbook,
				sheetIndex,
				r.top,
				Math.min(r.bottom, r.top + 100000),
				hidden,
			);
		}
	});
}

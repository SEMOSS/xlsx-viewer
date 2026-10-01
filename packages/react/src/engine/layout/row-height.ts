import type { EditTransaction, StyleResolver, Workbook, Worksheet } from 'xlsx-viewer-core';
import { cellKey, keyRow, MAX_COL_INDEX, setAutoRowHeight } from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import { describeCellContent } from '../render/cell-content';
import { canvasFont } from '../render/fonts';
import { measureLines } from '../render/text-measure';
import { CELL_PADDING_X } from './viewport-layout';

/** Rows re-fitted per edit at most (a huge paste keeps its heights). */
const MAX_AUTO_ROWS = 200;

/**
 * Height in points of one line of text at `sizePt`, matching Excel's
 * automatic row heights for common fonts within a pixel (11 pt gives 15 pt).
 */
export function lineHeightPt(sizePt: number): number {
	return Math.ceil(sizePt * 1.708 + 0.5) * 0.75;
}

/**
 * Height in points needed by the content of a row. Merged cells are
 * ignored, as in Excel; wrapped text counts its lines at the column width.
 */
export function measureRowContent(
	workbook: Workbook,
	styles: StyleResolver,
	sheet: Worksheet,
	row: number,
	geometry: SheetGeometry = new SheetGeometry(sheet, { zoom: 1 }),
): number {
	let maxPt = 0;
	for (let col = 0; col <= MAX_COL_INDEX; col++) {
		const cell = sheet.cells.get(cellKey(row, col));
		if (!cell) {
			if (col > 2000) {
				break;
			}
			continue;
		}
		if (
			sheet.merges.some((m) => row >= m.top && row <= m.bottom && col >= m.left && col <= m.right)
		) {
			continue;
		}
		const style = styles.resolve(cell.style);
		const sizePt = style.font.size ?? 11;
		let lines = 1;
		if (style.alignment.wrapText && cell.value !== null) {
			const content = describeCellContent(cell, style, {
				zoom: 1,
				colorCtx: { theme: workbook.theme },
				date1904: workbook.date1904,
				defaultTextColor: '#000',
				showFormulas: sheet.view.showFormulas,
			});
			if (content) {
				const font = canvasFont(content.font);
				const avail = Math.max(10, geometry.colWidth(col) - CELL_PADDING_X * 2 - content.indent);
				lines = 0;
				for (const para of content.text.split('\n')) {
					lines += Math.max(1, Math.ceil(measureLines(para, font) / avail));
				}
			}
		}
		maxPt = Math.max(maxPt, lineHeightPt(sizePt) * lines);
	}
	return maxPt;
}

/**
 * Excel's automatic row height: re-fit rows holding cells changed by the
 * transaction, unless the row has a custom (user-set) height or is hidden.
 * Runs inside the transaction so undo restores the old heights.
 */
export function autoFitTouchedRows(
	tx: EditTransaction,
	workbook: Workbook,
	styles: StyleResolver,
): void {
	for (const sheetIndex of tx.touchedSheetIndices()) {
		const sheet = workbook.sheets[sheetIndex];
		if (!sheet) {
			continue;
		}
		const rows = new Set<number>();
		for (const key of tx.touchedCellKeys(sheetIndex)) {
			rows.add(keyRow(key));
			if (rows.size > MAX_AUTO_ROWS) {
				break;
			}
		}
		if (rows.size === 0 || rows.size > MAX_AUTO_ROWS) {
			continue;
		}
		const geometry = new SheetGeometry(sheet, { zoom: 1 });
		for (const row of rows) {
			const info = sheet.rows.get(row);
			if (info?.customHeight || info?.hidden) {
				continue;
			}
			const pt = measureRowContent(workbook, styles, sheet, row, geometry);
			const target = pt > sheet.defaultRowHeight + 0.01 ? pt : undefined;
			// Keep heights Excel computed when ours is within 2 px (font metrics differ slightly).
			const current = info?.height ?? sheet.defaultRowHeight;
			if (Math.abs(current - (target ?? sheet.defaultRowHeight)) > 1.5) {
				setAutoRowHeight(tx, workbook, sheetIndex, row, target ?? null);
			}
		}
	}
}

import { updateCols, updateRow } from '../sheet/worksheet';
import type { CellAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/** Set the stored width (character units) of columns `min..max`. */
export function setColumnWidth(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	min: number,
	max: number,
	width: number,
): void {
	tx.touchSheet(sheetIndex, false, 'formatColumns');
	const sheet = workbook.sheets[sheetIndex];
	updateCols(sheet, min, max, (info) => ({
		...info,
		width: width > 0 ? width : info.width,
		customWidth: width > 0 ? true : info.customWidth,
		hidden: width <= 0 ? true : undefined,
		bestFit: undefined,
	}));
}

/** Set row heights in points; `null` resets to automatic height. */
export function setRowHeight(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	top: number,
	bottom: number,
	heightPt: number | null,
): void {
	tx.touchRows(sheetIndex, top, bottom);
	const sheet = workbook.sheets[sheetIndex];
	for (let row = top; row <= bottom; row++) {
		updateRow(sheet, row, (info) =>
			heightPt === null
				? { ...info, height: undefined, customHeight: undefined }
				: heightPt <= 0
					? { ...info, hidden: true }
					: { ...info, height: heightPt, customHeight: true, hidden: undefined },
		);
	}
}

/**
 * Automatic row height: Excel re-fits rows without a custom height to their
 * content. Stores the height without `customHeight`; `null` clears it.
 */
export function setAutoRowHeight(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	heightPt: number | null,
): void {
	tx.touchRows(sheetIndex, row, row);
	updateRow(workbook.sheets[sheetIndex], row, (info) => ({
		...info,
		height: heightPt ?? undefined,
		customHeight: undefined,
	}));
}

export function setRowsHidden(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	top: number,
	bottom: number,
	hidden: boolean,
): void {
	tx.touchRows(sheetIndex, top, bottom);
	const sheet = workbook.sheets[sheetIndex];
	for (let row = top; row <= bottom; row++) {
		updateRow(sheet, row, (info) => ({ ...info, hidden: hidden || undefined }));
	}
}

export function setColumnsHidden(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	min: number,
	max: number,
	hidden: boolean,
): void {
	tx.touchSheet(sheetIndex, false, 'formatColumns');
	const sheet = workbook.sheets[sheetIndex];
	updateCols(sheet, min, max, (info) => ({
		...info,
		hidden: hidden || undefined,
		// Unhiding a column hidden by zero width restores the default width.
		width: !hidden && info.width === 0 ? undefined : info.width,
	}));
}

/**
 * Freeze panes so that rows above and columns left of `at` stay visible
 * (View > Freeze Panes). `null` unfreezes.
 */
export function setFreezePanes(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	at: CellAddress | null,
	topLeftVisible?: CellAddress,
): void {
	tx.touchSheet(sheetIndex, false, 'view');
	const sheet = workbook.sheets[sheetIndex];
	if (!at || (at.row === 0 && at.col === 0)) {
		sheet.view = { ...sheet.view, pane: undefined };
	} else {
		const origin = topLeftVisible ?? { row: 0, col: 0 };
		sheet.view = {
			...sheet.view,
			pane: {
				xSplit: Math.max(0, at.col - origin.col),
				ySplit: Math.max(0, at.row - origin.row),
				topLeftCell: at,
				activePane:
					at.row > 0 && at.col > 0 ? 'bottomRight' : at.row > 0 ? 'bottomLeft' : 'topRight',
				state: 'frozen',
			},
			topLeftCell: origin,
		};
	}
	sheet.dirty = true;
}

/** Set the sheet zoom in percent (10-400). */
export function setSheetZoom(workbook: Workbook, sheetIndex: number, zoomPercent: number): void {
	const sheet = workbook.sheets[sheetIndex];
	const z = Math.max(10, Math.min(400, Math.round(zoomPercent)));
	if (sheet.view.zoomScale !== z) {
		sheet.view = { ...sheet.view, zoomScale: z };
		sheet.dirty = true;
	}
}

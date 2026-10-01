import { getColInfo } from '../sheet/worksheet';
import { columnWidthToPixels, defaultColumnWidthPx } from '../styles/column-width';
import { DEFAULT_MAX_DIGIT_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from '../styles/defaults';
import type { Worksheet } from '../types/sheet';

/**
 * Row heights and column widths as printed, in points at 100%. Rows use
 * their height in points; columns convert their stored width to pixels at
 * the 7 px maximum digit width of Calibri 11 (96 DPI), then to points.
 * Hidden rows and columns measure 0. The rules match the screen geometry,
 * so printed and displayed cells line up.
 */

const PX_TO_PT = 0.75;

/** Default row height of the sheet in points (0 when rows are hidden by default). */
export function defaultPrintRowHeight(sheet: Worksheet): number {
	return sheet.zeroHeight ? 0 : sheet.defaultRowHeight || DEFAULT_ROW_HEIGHT_PT;
}

/** Printed height of a row in points (0 when hidden). */
export function printRowHeight(sheet: Worksheet, row: number): number {
	const info = sheet.rows.get(row);
	if (!info) {
		return defaultPrintRowHeight(sheet);
	}
	if (info.hidden) {
		return 0;
	}
	if (info.height !== undefined) {
		return Math.max(0, info.height);
	}
	return sheet.defaultRowHeight || DEFAULT_ROW_HEIGHT_PT;
}

/** Default column width of the sheet in points. */
export function defaultPrintColWidth(
	sheet: Worksheet,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	return defaultColumnWidthPx(sheet.defaultColWidth, sheet.baseColWidth, maxDigitWidth) * PX_TO_PT;
}

/** Printed width of a column in points (0 when hidden). */
export function printColWidth(
	sheet: Worksheet,
	col: number,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	const info = getColInfo(sheet, col);
	if (info?.hidden) {
		return 0;
	}
	if (info?.width !== undefined) {
		return columnWidthToPixels(info.width, maxDigitWidth) * PX_TO_PT;
	}
	return defaultPrintColWidth(sheet, maxDigitWidth);
}

/** Size of one row or column in points. */
export type AxisSize = (index: number) => number;

/** Size functions of both axes for a sheet. */
export function printAxes(sheet: Worksheet): { row: AxisSize; col: AxisSize } {
	return {
		row: (r) => printRowHeight(sheet, r),
		col: (c) => printColWidth(sheet, c),
	};
}

/** Sum of sizes over an inclusive span. */
export function spanSize(size: AxisSize, start: number, end: number): number {
	let total = 0;
	for (let i = start; i <= end; i++) {
		total += size(i);
	}
	return total;
}

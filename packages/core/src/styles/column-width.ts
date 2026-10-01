import {
	DEFAULT_COL_WIDTH_CHARS,
	DEFAULT_MAX_DIGIT_WIDTH_PX,
	DEFAULT_ROW_HEIGHT_PT,
} from './defaults';

/**
 * Column widths are stored in "characters": multiples of the maximum digit
 * width of the workbook's default font, including cell padding. These
 * conversions follow ECMA-376 Part 1, 18.3.1.13 at 96 DPI.
 */

/** Stored width (characters) to pixels. */
export function columnWidthToPixels(
	width: number,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	if (width <= 0) {
		return 0;
	}
	return Math.trunc(((256 * width + Math.trunc(128 / maxDigitWidth)) / 256) * maxDigitWidth);
}

/** Pixels to stored width (characters), rounded to 1/256 like Excel. */
export function pixelsToColumnWidth(
	px: number,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	if (px <= 0) {
		return 0;
	}
	return Math.trunc((px / maxDigitWidth) * 256) / 256;
}

/** Visible characters shown in Excel's "Column Width" dialog for a stored width. */
export function storedWidthToDisplayChars(
	width: number,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	const px = columnWidthToPixels(width, maxDigitWidth);
	return Math.max(0, Math.round(((px - 5) / maxDigitWidth) * 100) / 100);
}

/** Stored width for a number of visible characters typed in the "Column Width" dialog. */
export function displayCharsToStoredWidth(
	chars: number,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	if (chars <= 0) {
		return 0;
	}
	return Math.trunc(((chars * maxDigitWidth + 5) / maxDigitWidth) * 256) / 256;
}

/**
 * Default column width in pixels from `sheetFormatPr`: `defaultColWidth`
 * when present, else `baseColWidth` characters plus 5px padding.
 */
export function defaultColumnWidthPx(
	defaultColWidth: number | undefined,
	baseColWidth: number | undefined,
	maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
	if (defaultColWidth !== undefined && defaultColWidth > 0) {
		return columnWidthToPixels(defaultColWidth, maxDigitWidth);
	}
	if (baseColWidth !== undefined) {
		// Excel rounds the padded base width up to a multiple of 8 pixels.
		const px = baseColWidth * maxDigitWidth + 5;
		return Math.ceil(px / 8) * 8;
	}
	return columnWidthToPixels(DEFAULT_COL_WIDTH_CHARS, maxDigitWidth);
}

/** Points to pixels at 96 DPI. */
export function pointsToPixels(pt: number): number {
	return (pt * 96) / 72;
}

/** Pixels to points at 96 DPI. */
export function pixelsToPoints(px: number): number {
	return (px * 72) / 96;
}

/** Default row height in pixels. */
export function defaultRowHeightPx(defaultRowHeightPt = DEFAULT_ROW_HEIGHT_PT): number {
	return Math.round(pointsToPixels(defaultRowHeightPt));
}

/** EMU per pixel at 96 DPI. */
export const EMU_PER_PIXEL = 9525;
export const EMU_PER_POINT = 12700;
export const EMU_PER_INCH = 914400;

export function emuToPixels(emu: number): number {
	return emu / EMU_PER_PIXEL;
}

export function pixelsToEmu(px: number): number {
	return Math.round(px * EMU_PER_PIXEL);
}

import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { rangeContains, rangesIntersect } from '../address/range';
import { subtractRange } from '../conditional/edit';
import { formatValue } from '../numfmt/format-value';
import { findMerge } from '../sheet/worksheet';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Hyperlink, Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { findHyperlinkCellStyle, hyperlinkCellXf } from './hyperlink-style';
import type { EditTransaction } from './transaction';

/**
 * Insert > Link (Ctrl+K): add, replace and remove cell hyperlinks. External
 * targets are written as relationships, `location` stays inside the file.
 */

export {
	buildMailtoAddress,
	formatHyperlinkLocation,
	type HyperlinkLocation,
	normalizeHyperlinkAddress,
	parseHyperlinkLocation,
	parseMailtoAddress,
} from './hyperlink-address';
export {
	ensureHyperlinkCellStyle,
	findHyperlinkCellStyle,
	HYPERLINK_BUILTIN_STYLE_ID,
	hyperlinkCellXf,
} from './hyperlink-style';

/** What a hyperlink points at: an external `target`, a `location` in the workbook, or both. */
export interface HyperlinkSpec {
	target?: string;
	location?: string;
	/** ScreenTip text. */
	tooltip?: string;
}

/** Cells styled with the Hyperlink style at most; larger ranges only style their existing cells. */
const MAX_STYLED_CELLS = 10_000;

/** The hyperlink covering a cell (the most recently added wins), if any. */
export function hyperlinkAt(sheet: Worksheet, row: number, col: number): Hyperlink | undefined {
	for (let i = sheet.hyperlinks.length - 1; i >= 0; i--) {
		if (rangeContains(sheet.hyperlinks[i].ref, row, col)) {
			return sheet.hyperlinks[i];
		}
	}
	return undefined;
}

/** Displayed text of a cell (value through its number format), as Text to display shows it. */
export function hyperlinkDisplayText(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): string {
	const cell = workbook.sheets[sheetIndex]?.cells.get(cellKey(row, col));
	if (!cell || cell.value === null) {
		return '';
	}
	const styles = workbook.styles;
	const code = numberFormatCode(styles.cellXfs[cell.style ?? 0]?.numFmtId ?? 0, styles.numFmts);
	return formatValue(cell.value, code, { date1904: workbook.date1904 }).text;
}

/** Links with `range` cut out: overlapping links keep only the parts outside it. */
function cutLinks(links: readonly Hyperlink[], range: RangeAddress): Hyperlink[] {
	const out: Hyperlink[] = [];
	for (const link of links) {
		if (!rangesIntersect(link.ref, range)) {
			out.push(link);
			continue;
		}
		for (const part of subtractRange(link.ref, range)) {
			out.push({ ...link, ref: part });
		}
	}
	return out;
}

/** Keys of the cells of `range`, or only of its existing cells when the range is large. */
function cellsOf(sheet: Worksheet, range: RangeAddress, all: boolean): number[] {
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	if (all && area <= MAX_STYLED_CELLS) {
		const keys: number[] = [];
		for (let row = range.top; row <= range.bottom; row++) {
			for (let col = range.left; col <= range.right; col++) {
				keys.push(cellKey(row, col));
			}
		}
		return keys;
	}
	const keys = [...sheet.cells.keys()].filter((k) => rangeContains(range, keyRow(k), keyCol(k)));
	const topLeft = cellKey(range.top, range.left);
	return all && !keys.includes(topLeft) ? [topLeft, ...keys] : keys;
}

function hiddenByMerge(sheet: Worksheet, key: number): boolean {
	const row = keyRow(key);
	const col = keyCol(key);
	const merge = findMerge(sheet, row, col);
	return merge !== undefined && (merge.top !== row || merge.left !== col);
}

function writeDisplayText(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	text: string,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const { top: row, left: col } = range;
	const cell = sheet.cells.get(cellKey(row, col));
	const blank = !cell || (cell.value === null && !cell.formula);
	if (
		!blank &&
		(cell.value === text || hyperlinkDisplayText(workbook, sheetIndex, row, col) === text)
	) {
		return;
	}
	tx.touchCell(sheetIndex, row, col);
	// Plain text (no quote prefix), so "123" stays text; the cell keeps its format.
	const next: Cell = { value: text };
	if (cell?.style) {
		next.style = cell.style;
	}
	sheet.cells.set(cellKey(row, col), next);
}

/**
 * Link the cells of `range` (Insert Hyperlink OK): links overlapping the
 * range are replaced, `displayText` (when given) goes into the top-left
 * cell if it shows something else, and unformatted cells of the range get
 * the Hyperlink cell style (hidden cells of merged areas are skipped).
 */
export function setHyperlink(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	link: HyperlinkSpec,
	displayText?: string,
): Hyperlink {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchSheet(sheetIndex);
	tx.touchRange(sheetIndex, range, false);
	const hyperlink: Hyperlink = { ref: { ...range } };
	if (link.target) {
		hyperlink.target = link.target;
	}
	if (link.location) {
		hyperlink.location = link.location;
	}
	if (link.tooltip) {
		hyperlink.tooltip = link.tooltip;
	}
	if (displayText) {
		hyperlink.display = displayText;
		writeDisplayText(tx, workbook, sheetIndex, range, displayText);
	}
	sheet.hyperlinks = [...cutLinks(sheet.hyperlinks, range), hyperlink];
	const xf = hyperlinkCellXf(workbook.styles);
	for (const key of cellsOf(sheet, range, true)) {
		const cell = sheet.cells.get(key);
		if (!cell?.style && !hiddenByMerge(sheet, key)) {
			sheet.cells.set(key, cell ? { ...cell, style: xf } : { value: null, style: xf });
		}
	}
	sheet.dirty = true;
	return hyperlink;
}

/**
 * Remove Link / Remove Hyperlinks: unlink the cells of `ranges` (links
 * reaching outside keep their other cells) and return cells formatted with
 * the Hyperlink style to the Normal style, keeping their values. Returns
 * the number of links that were affected.
 */
export function removeHyperlinks(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
): number {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchSheet(sheetIndex);
	const affected = sheet.hyperlinks.filter((h) => ranges.some((r) => rangesIntersect(h.ref, r)));
	let links: Hyperlink[] = sheet.hyperlinks;
	for (const range of ranges) {
		links = cutLinks(links, range);
	}
	sheet.hyperlinks = links;
	const styleXf = findHyperlinkCellStyle(workbook.styles);
	if (styleXf !== undefined) {
		const xfs = workbook.styles.cellXfs;
		for (const range of ranges) {
			for (const key of cellsOf(sheet, range, false)) {
				const cell = sheet.cells.get(key);
				if (!cell?.style || xfs[cell.style]?.xfId !== styleXf) {
					continue;
				}
				const { style: _style, ...rest } = cell;
				if (rest.value === null && !rest.formula) {
					sheet.cells.delete(key);
				} else {
					sheet.cells.set(key, rest);
				}
			}
		}
	}
	sheet.dirty = true;
	return affected.length;
}

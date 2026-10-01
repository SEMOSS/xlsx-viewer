import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { rangeContains, rangesIntersect } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

export type ClearKind = 'all' | 'contents' | 'formats' | 'comments' | 'hyperlinks';

function forEachExistingCell(
	cells: Map<number, unknown>,
	range: RangeAddress,
	visit: (key: number, row: number, col: number) => void,
): void {
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	if (area > cells.size) {
		for (const key of [...cells.keys()]) {
			const row = keyRow(key);
			const col = keyCol(key);
			if (rangeContains(range, row, col)) {
				visit(key, row, col);
			}
		}
		return;
	}
	for (let row = range.top; row <= range.bottom; row++) {
		for (let col = range.left; col <= range.right; col++) {
			const key = cellKey(row, col);
			if (cells.has(key)) {
				visit(key, row, col);
			}
		}
	}
}

/**
 * Home > Clear: remove contents (values and formulas), formats, comments
 * and notes, hyperlinks, or everything, from the given ranges.
 */
export function clearRanges(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
	kind: ClearKind = 'contents',
): void {
	const sheet = workbook.sheets[sheetIndex];
	for (const range of ranges) {
		if (kind === 'contents' || kind === 'formats' || kind === 'all') {
			tx.touchRange(sheetIndex, range, kind !== 'formats');
			forEachExistingCell(sheet.cells, range, (key) => {
				const cell = sheet.cells.get(key)!;
				if (kind === 'contents') {
					if (cell.style) {
						sheet.cells.set(key, { value: null, style: cell.style });
					} else {
						sheet.cells.delete(key);
					}
				} else if (kind === 'formats') {
					const { style: _s, richText, ...rest } = cell;
					if (rest.value === null && !rest.formula) {
						sheet.cells.delete(key);
					} else {
						sheet.cells.set(key, richText ? { ...rest } : rest);
					}
				} else {
					sheet.cells.delete(key);
				}
			});
		}
		if (kind === 'formats' || kind === 'all') {
			tx.touchSheet(sheetIndex);
			sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, range));
		}
		if (kind === 'comments' || kind === 'all') {
			tx.touchSheet(sheetIndex);
			const before = sheet.notes.length + sheet.threadedComments.length;
			sheet.notes = sheet.notes.filter((n) => !rangeContains(range, n.ref.row, n.ref.col));
			sheet.threadedComments = sheet.threadedComments.filter(
				(c) => !rangeContains(range, c.ref.row, c.ref.col),
			);
			if (sheet.notes.length + sheet.threadedComments.length !== before) {
				sheet.commentsDirty = true;
			}
		}
		if (kind === 'hyperlinks' || kind === 'all') {
			tx.touchSheet(sheetIndex);
			sheet.hyperlinks = sheet.hyperlinks.filter((h) => !rangesIntersect(h.ref, range));
		}
	}
	sheet.dirty = true;
}

import { cellKey } from '../address/cell-address';
import { rangesIntersect } from '../address/range';
import { StylePool } from '../styles/style-pool';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/** Ranges in `ranges` would discard values other than the top-left one (Excel warns). */
export function mergeWouldLoseData(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
): boolean {
	const sheet = workbook.sheets[sheetIndex];
	let count = 0;
	for (let row = range.top; row <= range.bottom; row++) {
		for (let col = range.left; col <= range.right; col++) {
			const cell = sheet.cells.get(cellKey(row, col));
			if (cell && (cell.value !== null || cell.formula)) {
				count++;
				if (count > 1) {
					return true;
				}
			}
		}
	}
	return false;
}

/**
 * Merge a range (Home > Merge Cells). Existing merges that overlap are
 * absorbed. Only the upper-left value is kept, as in Excel. With `center`,
 * the merged cell is centered horizontally (Merge & Center). With `across`,
 * each row is merged separately (Merge Across).
 */
export function mergeRange(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	options: { center?: boolean; across?: boolean } = {},
): void {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchSheet(sheetIndex);
	const targets: RangeAddress[] = options.across
		? Array.from({ length: range.bottom - range.top + 1 }, (_, i) => ({
				top: range.top + i,
				bottom: range.top + i,
				left: range.left,
				right: range.right,
			}))
		: [range];
	for (const target of targets) {
		if (target.top === target.bottom && target.left === target.right) {
			continue;
		}
		// Absorb overlapping merges into the new one.
		let merged = target;
		sheet.merges = sheet.merges.filter((m) => {
			if (rangesIntersect(m, merged)) {
				merged = {
					top: Math.min(m.top, merged.top),
					left: Math.min(m.left, merged.left),
					bottom: Math.max(m.bottom, merged.bottom),
					right: Math.max(m.right, merged.right),
				};
				return false;
			}
			return true;
		});
		for (let row = merged.top; row <= merged.bottom; row++) {
			for (let col = merged.left; col <= merged.right; col++) {
				if (row === merged.top && col === merged.left) {
					continue;
				}
				const key = cellKey(row, col);
				const cell = sheet.cells.get(key);
				if (cell && (cell.value !== null || cell.formula)) {
					const { formula: _f, array: _a, ...rest } = cell;
					sheet.cells.set(key, { ...rest, value: null });
					tx.markChanged({ sheet: sheetIndex, row, col });
				}
			}
		}
		sheet.merges.push(merged);
		if (options.center) {
			const pool = new StylePool(workbook.styles);
			const key = cellKey(merged.top, merged.left);
			const cell = sheet.cells.get(key);
			const style = pool.derive(cell?.style, { alignment: { horizontal: 'center' } });
			sheet.cells.set(key, cell ? { ...cell, style } : { value: null, style });
		}
	}
	sheet.dirty = true;
}

/** Unmerge every merged area intersecting `range`. */
export function unmergeRange(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
): void {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchSheet(sheetIndex);
	sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, range));
	sheet.dirty = true;
}

/** Expand a range so it fully contains every merged area it touches (Excel's selection rule). */
export function expandToMerges(merges: readonly RangeAddress[], range: RangeAddress): RangeAddress {
	let r = { ...range };
	let changed = true;
	while (changed) {
		changed = false;
		for (const m of merges) {
			if (rangesIntersect(m, r)) {
				const next = {
					top: Math.min(r.top, m.top),
					left: Math.min(r.left, m.left),
					bottom: Math.max(r.bottom, m.bottom),
					right: Math.max(r.right, m.right),
				};
				if (
					next.top !== r.top ||
					next.left !== r.left ||
					next.bottom !== r.bottom ||
					next.right !== r.right
				) {
					r = next;
					changed = true;
				}
			}
		}
	}
	return r;
}

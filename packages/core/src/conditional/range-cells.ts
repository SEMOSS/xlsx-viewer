import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { rangesIntersect } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { CellValue } from '../types/cell';
import type { Worksheet } from '../types/sheet';

/** Normalised copy of a range clamped to the grid, or `undefined` when unusable. */
export function sanitizeRange(r: RangeAddress): RangeAddress | undefined {
	if (![r.top, r.left, r.bottom, r.right].every(Number.isFinite)) {
		return undefined;
	}
	const top = Math.max(0, Math.min(r.top, r.bottom));
	const bottom = Math.min(MAX_ROW_INDEX, Math.max(r.top, r.bottom));
	const left = Math.max(0, Math.min(r.left, r.right));
	const right = Math.min(MAX_COL_INDEX, Math.max(r.left, r.right));
	return top <= bottom && left <= right ? { top, left, bottom, right } : undefined;
}

/** Smallest range covering every input range, or `undefined` for an empty list. */
export function boundingBox(ranges: readonly RangeAddress[]): RangeAddress | undefined {
	if (ranges.length === 0) {
		return undefined;
	}
	let { top, left, bottom, right } = ranges[0];
	for (let i = 1; i < ranges.length; i++) {
		const r = ranges[i];
		top = Math.min(top, r.top);
		left = Math.min(left, r.left);
		bottom = Math.max(bottom, r.bottom);
		right = Math.max(right, r.right);
	}
	return { top, left, bottom, right };
}

export function anyRangeContains(
	ranges: readonly RangeAddress[],
	row: number,
	col: number,
): boolean {
	for (const r of ranges) {
		if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
			return true;
		}
	}
	return false;
}

function hasOverlap(ranges: readonly RangeAddress[]): boolean {
	if (ranges.length > 64) {
		return true;
	}
	for (let i = 0; i < ranges.length; i++) {
		for (let j = i + 1; j < ranges.length; j++) {
			if (rangesIntersect(ranges[i], ranges[j])) {
				return true;
			}
		}
	}
	return false;
}

/**
 * Visit every stored cell inside the ranges once (blank cells without an
 * entry are skipped). Picks the cheaper strategy: probing each address of
 * small ranges, or scanning the sheet's cell map for large ones such as
 * whole-column references. Visit order is unspecified.
 */
export function forEachCellInRanges(
	sheet: Worksheet,
	ranges: readonly RangeAddress[],
	visit: (value: CellValue, row: number, col: number) => void,
): void {
	const cells = sheet.cells;
	if (cells.size === 0 || ranges.length === 0) {
		return;
	}
	let area = 0;
	for (const r of ranges) {
		area += (r.bottom - r.top + 1) * (r.right - r.left + 1);
	}
	if (area <= cells.size * 2) {
		const seen = ranges.length > 1 && hasOverlap(ranges) ? new Set<number>() : undefined;
		for (const r of ranges) {
			for (let row = r.top; row <= r.bottom; row++) {
				for (let col = r.left; col <= r.right; col++) {
					const key = cellKey(row, col);
					if (seen) {
						if (seen.has(key)) {
							continue;
						}
						seen.add(key);
					}
					const cell = cells.get(key);
					if (cell) {
						visit(cell.value, row, col);
					}
				}
			}
		}
		return;
	}
	const box = boundingBox(ranges)!;
	for (const [key, cell] of cells) {
		const row = keyRow(key);
		if (row < box.top || row > box.bottom) {
			continue;
		}
		const col = keyCol(key);
		if (col < box.left || col > box.right) {
			continue;
		}
		if (ranges.length === 1 || anyRangeContains(ranges, row, col)) {
			visit(cell.value, row, col);
		}
	}
}

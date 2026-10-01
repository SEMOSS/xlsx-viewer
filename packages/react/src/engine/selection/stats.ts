import type { RangeAddress, Worksheet } from 'xlsx-viewer-core';
import { cellKey, keyCol, keyRow } from 'xlsx-viewer-core';

export interface SelectionStats {
	/** Non-empty cells. */
	count: number;
	/** Numeric cells. */
	numericCount: number;
	sum: number;
	min: number | null;
	max: number | null;
	average: number | null;
}

/**
 * Status-bar statistics of the selected ranges. Iterates stored cells for
 * large areas so whole-column selections stay fast; overlapping ranges are
 * counted once. Hidden rows are skipped, like Excel.
 */
export function selectionStats(sheet: Worksheet, ranges: readonly RangeAddress[]): SelectionStats {
	const seen = new Set<number>();
	let count = 0;
	let numericCount = 0;
	let sum = 0;
	let min: number | null = null;
	let max: number | null = null;
	const visit = (key: number) => {
		if (seen.has(key)) {
			return;
		}
		seen.add(key);
		const cell = sheet.cells.get(key);
		if (!cell || cell.value === null || cell.value === '') {
			return;
		}
		if (sheet.rows.get(keyRow(key))?.hidden) {
			return;
		}
		count++;
		if (typeof cell.value === 'number') {
			numericCount++;
			sum += cell.value;
			min = min === null ? cell.value : Math.min(min, cell.value);
			max = max === null ? cell.value : Math.max(max, cell.value);
		}
	};
	for (const r of ranges) {
		const area = (r.bottom - r.top + 1) * (r.right - r.left + 1);
		if (area > sheet.cells.size * 2) {
			for (const key of sheet.cells.keys()) {
				const row = keyRow(key);
				const col = keyCol(key);
				if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
					visit(key);
				}
			}
		} else {
			for (let row = r.top; row <= r.bottom; row++) {
				for (let col = r.left; col <= r.right; col++) {
					visit(cellKey(row, col));
				}
			}
		}
	}
	return {
		count,
		numericCount,
		sum,
		min,
		max,
		average: numericCount > 0 ? sum / numericCount : null,
	};
}

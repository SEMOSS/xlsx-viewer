import type { Cell, Worksheet } from 'xlsx-viewer-core';
import { cellKey, getColInfo } from 'xlsx-viewer-core';

/**
 * Effective cellXfs index for a position: the cell's own style, else the
 * row's default style (when `customFormat`), else the column's style.
 */
export function styleIdAt(sheet: Worksheet, row: number, col: number, cell?: Cell): number {
	const c = cell ?? sheet.cells.get(cellKey(row, col));
	if (c && c.style !== undefined) {
		return c.style;
	}
	const r = sheet.rows.get(row);
	if (r?.customFormat && r.style !== undefined) {
		return r.style;
	}
	return getColInfo(sheet, col)?.style ?? 0;
}

/** Index of merged ranges by their cells, for fast lookups while laying out a viewport. */
export class MergeIndex {
	private readonly byCell = new Map<number, number>();

	constructor(
		readonly merges: ReadonlyArray<{ top: number; left: number; bottom: number; right: number }>,
	) {
		merges.forEach((m, i) => {
			const area = (m.bottom - m.top + 1) * (m.right - m.left + 1);
			if (area > 50000) {
				return; // Huge merges are resolved by scanning (rare).
			}
			for (let r = m.top; r <= m.bottom; r++) {
				for (let c = m.left; c <= m.right; c++) {
					this.byCell.set(cellKey(r, c), i);
				}
			}
		});
	}

	/** Merge containing the cell, if any. */
	find(
		row: number,
		col: number,
	): { top: number; left: number; bottom: number; right: number } | undefined {
		const i = this.byCell.get(cellKey(row, col));
		if (i !== undefined) {
			return this.merges[i];
		}
		for (const m of this.merges) {
			const area = (m.bottom - m.top + 1) * (m.right - m.left + 1);
			if (area > 50000 && row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
				return m;
			}
		}
		return undefined;
	}
}

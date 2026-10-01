import type { UndoEntry, Workbook, Worksheet } from 'xlsx-viewer-core';
import { usedRange } from 'xlsx-viewer-core';

/**
 * Rows and columns spanned by content and formatting (the grid's scroll
 * extents). Computed once per sheet, then grown by edits instead of
 * rescanning every cell; like Excel's used range it only shrinks after a
 * structural change.
 */
export class UsedExtents {
	private extents = new WeakMap<Worksheet, { rows: number; cols: number }>();

	of(sheet: Worksheet): { rows: number; cols: number } {
		let extent = this.extents.get(sheet);
		if (!extent) {
			const used = usedRange(sheet, true);
			extent = { rows: used ? used.bottom + 1 : 0, cols: used ? used.right + 1 : 0 };
			this.extents.set(sheet, extent);
		}
		return extent;
	}

	/** Grow with the cells an edit changed; structural edits recompute. */
	track(entry: UndoEntry, workbook: Workbook): void {
		if (entry.structural || entry.fullRecalc) {
			for (const i of entry.sheets) {
				const sheet = workbook.sheets[i];
				if (sheet) {
					this.extents.delete(sheet);
				}
			}
			return;
		}
		for (const pos of entry.changedCells) {
			const sheet = workbook.sheets[pos.sheet];
			const extent = sheet && this.extents.get(sheet);
			if (extent) {
				extent.rows = Math.max(extent.rows, pos.row + 1);
				extent.cols = Math.max(extent.cols, pos.col + 1);
			}
		}
	}

	/** Forget everything (the workbook changed outside transactions). */
	reset(): void {
		this.extents = new WeakMap();
	}
}

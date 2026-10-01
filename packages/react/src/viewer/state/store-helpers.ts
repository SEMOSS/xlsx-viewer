import type { Workbook, Worksheet } from 'xlsx-viewer-core';
import { expandToMerges } from 'xlsx-viewer-core';

import { singleCellSelection, type SelectionState } from '../../engine/selection/selection';

/** The selection a sheet opens with (its saved selection, covering whole merged areas). */
export function initialSelection(sheet: Worksheet): SelectionState {
	const sel = sheet.view.selection;
	const base: SelectionState = sel
		? {
				active: sel.activeCell,
				anchor: sel.activeCell,
				ranges: sel.ranges.length
					? sel.ranges
					: [
							{
								top: sel.activeCell.row,
								left: sel.activeCell.col,
								bottom: sel.activeCell.row,
								right: sel.activeCell.col,
							},
						],
			}
		: singleCellSelection(0, 0);
	if (sheet.merges.length === 0) {
		return base;
	}
	// Selections always cover whole merged areas; the active cell is a merge's top-left.
	const merge = sheet.merges.find(
		(m) =>
			base.active.row >= m.top &&
			base.active.row <= m.bottom &&
			base.active.col >= m.left &&
			base.active.col <= m.right,
	);
	const active = merge ? { row: merge.top, col: merge.left } : base.active;
	return {
		active,
		anchor: active,
		ranges: base.ranges.map((r) => expandToMerges(sheet.merges, r)),
	};
}

/** Some formula needs computing before display (no cached value, stale, or a full calculation on load). */
export function needsCalculation(workbook: Workbook): boolean {
	if (workbook.calc.fullCalcOnLoad) {
		return true;
	}
	for (const sheet of workbook.sheets) {
		for (const cell of sheet.cells.values()) {
			if (cell.formula !== undefined && (cell.value === null || cell.stale)) {
				return true;
			}
		}
	}
	return false;
}

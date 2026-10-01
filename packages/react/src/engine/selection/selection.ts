import type { CellAddress, RangeAddress } from 'xlsx-viewer-core';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from 'xlsx-viewer-core';

/** Selection of one sheet: several ranges, the last being the current one. */
export interface SelectionState {
	/** The active cell (where typing goes). */
	active: CellAddress;
	/** Fixed corner of the current range (Shift+click / Shift+arrows extend from here). */
	anchor: CellAddress;
	ranges: RangeAddress[];
}

export function singleCellSelection(row: number, col: number): SelectionState {
	return {
		active: { row, col },
		anchor: { row, col },
		ranges: [{ top: row, left: col, bottom: row, right: col }],
	};
}

export function currentRange(sel: SelectionState): RangeAddress {
	return sel.ranges[sel.ranges.length - 1];
}

export function rangeFromCorners(a: CellAddress, b: CellAddress): RangeAddress {
	return {
		top: Math.min(a.row, b.row),
		left: Math.min(a.col, b.col),
		bottom: Math.max(a.row, b.row),
		right: Math.max(a.col, b.col),
	};
}

export function wholeColumns(left: number, right: number): RangeAddress {
	return {
		top: 0,
		bottom: MAX_ROW_INDEX,
		left: Math.min(left, right),
		right: Math.max(left, right),
	};
}

export function wholeRows(top: number, bottom: number): RangeAddress {
	return {
		left: 0,
		right: MAX_COL_INDEX,
		top: Math.min(top, bottom),
		bottom: Math.max(top, bottom),
	};
}

export function selectAll(): RangeAddress {
	return { top: 0, left: 0, bottom: MAX_ROW_INDEX, right: MAX_COL_INDEX };
}

/** Rows touched by any selected range (for header highlighting). */
export function selectionTouchesRow(sel: SelectionState, row: number): boolean {
	return sel.ranges.some((r) => row >= r.top && row <= r.bottom);
}

export function selectionTouchesCol(sel: SelectionState, col: number): boolean {
	return sel.ranges.some((r) => col >= r.left && col <= r.right);
}

/** A whole row is selected (header shown in the "fully selected" color). */
export function isRowFullySelected(sel: SelectionState, row: number): boolean {
	return sel.ranges.some(
		(r) => r.left === 0 && r.right === MAX_COL_INDEX && row >= r.top && row <= r.bottom,
	);
}

export function isColFullySelected(sel: SelectionState, col: number): boolean {
	return sel.ranges.some(
		(r) => r.top === 0 && r.bottom === MAX_ROW_INDEX && col >= r.left && col <= r.right,
	);
}

export function selectionCellCount(sel: SelectionState): number {
	let n = 0;
	for (const r of sel.ranges) {
		n += (r.bottom - r.top + 1) * (r.right - r.left + 1);
	}
	return n;
}

/**
 * Move the active cell inside a multi-cell selection (Enter / Tab with a
 * selection active): wraps across rows / columns and ranges like Excel.
 */
export function cycleActiveWithinSelection(
	sel: SelectionState,
	direction: 'down' | 'up' | 'right' | 'left',
): SelectionState {
	const idx = sel.ranges.findIndex(
		(r) =>
			sel.active.row >= r.top &&
			sel.active.row <= r.bottom &&
			sel.active.col >= r.left &&
			sel.active.col <= r.right,
	);
	let rangeIndex = idx < 0 ? sel.ranges.length - 1 : idx;
	let r = sel.ranges[rangeIndex];
	let { row, col } = sel.active;
	const nextRange = (step: 1 | -1) => {
		rangeIndex = (rangeIndex + step + sel.ranges.length) % sel.ranges.length;
		r = sel.ranges[rangeIndex];
	};
	switch (direction) {
		case 'down':
			row++;
			if (row > r.bottom) {
				row = r.top;
				col++;
				if (col > r.right) {
					nextRange(1);
					row = r.top;
					col = r.left;
				}
			}
			break;
		case 'up':
			row--;
			if (row < r.top) {
				row = r.bottom;
				col--;
				if (col < r.left) {
					nextRange(-1);
					row = r.bottom;
					col = r.right;
				}
			}
			break;
		case 'right':
			col++;
			if (col > r.right) {
				col = r.left;
				row++;
				if (row > r.bottom) {
					nextRange(1);
					row = r.top;
					col = r.left;
				}
			}
			break;
		case 'left':
			col--;
			if (col < r.left) {
				col = r.right;
				row--;
				if (row < r.top) {
					nextRange(-1);
					row = r.bottom;
					col = r.right;
				}
			}
			break;
	}
	return { ...sel, active: { row, col } };
}

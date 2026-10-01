import { describe, expect, it } from 'vitest';
import type { RangeAddress } from 'xlsx-viewer-core';
import { MAX_COL_INDEX, MAX_ROW_INDEX, parseRange } from 'xlsx-viewer-core';

import {
	currentRange,
	cycleActiveWithinSelection,
	isColFullySelected,
	isRowFullySelected,
	rangeFromCorners,
	selectAll,
	selectionCellCount,
	selectionTouchesCol,
	selectionTouchesRow,
	singleCellSelection,
	wholeColumns,
	wholeRows,
	type SelectionState,
} from './selection';

const r = (ref: string): RangeAddress => parseRange(ref)!;

function selection(active: string, ...refs: string[]): SelectionState {
	const cell = r(active);
	return {
		active: { row: cell.top, col: cell.left },
		anchor: { row: cell.top, col: cell.left },
		ranges: refs.map(r),
	};
}

/** Active cells visited by repeated moves, as A1 references. */
function walk(
	sel: SelectionState,
	direction: 'down' | 'up' | 'right' | 'left',
	steps: number,
): string[] {
	const out: string[] = [];
	let s = sel;
	for (let i = 0; i < steps; i++) {
		s = cycleActiveWithinSelection(s, direction);
		out.push(`${String.fromCharCode(65 + s.active.col)}${s.active.row + 1}`);
	}
	return out;
}

describe('selection ranges', () => {
	it('starts as a single cell that is its own anchor and range', () => {
		expect(singleCellSelection(2, 3)).toEqual({
			active: { row: 2, col: 3 },
			anchor: { row: 2, col: 3 },
			ranges: [{ top: 2, left: 3, bottom: 2, right: 3 }],
		});
	});

	it('treats the last range as the current one', () => {
		expect(currentRange(selection('A1', 'A1:B2', 'D4:E5'))).toEqual(r('D4:E5'));
	});

	it('normalises corners in any order', () => {
		expect(rangeFromCorners({ row: 5, col: 1 }, { row: 2, col: 4 })).toEqual({
			top: 2,
			left: 1,
			bottom: 5,
			right: 4,
		});
		expect(rangeFromCorners({ row: 3, col: 3 }, { row: 3, col: 3 })).toEqual(r('D4'));
	});

	it('builds whole columns, whole rows and select-all', () => {
		expect(wholeColumns(3, 1)).toEqual({ top: 0, bottom: MAX_ROW_INDEX, left: 1, right: 3 });
		expect(wholeRows(4, 2)).toEqual({ top: 2, bottom: 4, left: 0, right: MAX_COL_INDEX });
		expect(selectAll()).toEqual({ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: MAX_COL_INDEX });
		expect(wholeColumns(1, 3)).toEqual(r('B:D'));
		expect(wholeRows(2, 4)).toEqual(r('3:5'));
	});
});

describe('header highlighting', () => {
	const sel = selection('B2', 'B2:C3', 'E7');

	it('reports rows and columns touched by any range', () => {
		expect([0, 1, 2, 3, 6].map((row) => selectionTouchesRow(sel, row))).toEqual([
			false,
			true,
			true,
			false,
			true,
		]);
		expect([0, 1, 2, 3, 4].map((col) => selectionTouchesCol(sel, col))).toEqual([
			false,
			true,
			true,
			false,
			true,
		]);
	});

	it('marks rows and columns as fully selected only for whole-line ranges', () => {
		expect(isRowFullySelected(sel, 1)).toBe(false);
		const rows = { ...sel, ranges: [...sel.ranges, wholeRows(9, 10)] };
		expect([8, 9, 10, 11].map((row) => isRowFullySelected(rows, row))).toEqual([
			false,
			true,
			true,
			false,
		]);
		expect(isColFullySelected(rows, 0)).toBe(false);
		const cols = { ...sel, ranges: [wholeColumns(2, 2)] };
		expect([1, 2, 3].map((col) => isColFullySelected(cols, col))).toEqual([false, true, false]);
		const all = { ...sel, ranges: [selectAll()] };
		expect(isRowFullySelected(all, 500) && isColFullySelected(all, 500)).toBe(true);
	});
});

describe('selectionCellCount', () => {
	it('adds up the areas of all ranges', () => {
		expect(selectionCellCount(singleCellSelection(0, 0))).toBe(1);
		expect(selectionCellCount(selection('A1', 'A1:B3', 'D1:D4'))).toBe(10);
		// Overlapping ranges are counted twice, like the number of cells Excel reports.
		expect(selectionCellCount(selection('A1', 'A1:B2', 'B2:C3'))).toBe(8);
		expect(selectionCellCount({ ...singleCellSelection(0, 0), ranges: [selectAll()] })).toBe(
			1048576 * 16384,
		);
	});
});

describe('cycleActiveWithinSelection', () => {
	it('moves down a column, then to the top of the next, then wraps to the start', () => {
		expect(walk(selection('A1', 'A1:B2'), 'down', 5)).toEqual(['A2', 'B1', 'B2', 'A1', 'A2']);
	});

	it('moves right along a row, then to the next row, then wraps', () => {
		expect(walk(selection('A1', 'A1:B2'), 'right', 5)).toEqual(['B1', 'A2', 'B2', 'A1', 'B1']);
	});

	it('moves up and left backwards through the range', () => {
		expect(walk(selection('A1', 'A1:B2'), 'up', 4)).toEqual(['B2', 'B1', 'A2', 'A1']);
		expect(walk(selection('A1', 'A1:B2'), 'left', 4)).toEqual(['B2', 'A2', 'B1', 'A1']);
	});

	it('continues into the next range and back to the first', () => {
		const sel = selection('A1', 'A1:A2', 'C3:D3');
		expect(walk(sel, 'down', 5)).toEqual(['A2', 'C3', 'D3', 'A1', 'A2']);
		expect(walk(sel, 'right', 3)).toEqual(['A2', 'C3', 'D3']);
		expect(walk(selection('C3', 'A1:A2', 'C3:D3'), 'up', 2)).toEqual(['A2', 'A1']);
		expect(walk(selection('A1', 'A1:A2', 'C3:D3'), 'left', 2)).toEqual(['D3', 'C3']);
	});

	it('keeps the anchor and ranges', () => {
		const sel = selection('B2', 'A1:C3');
		const next = cycleActiveWithinSelection(sel, 'right');
		expect(next.anchor).toEqual(sel.anchor);
		expect(next.ranges).toBe(sel.ranges);
		expect(next.active).toEqual({ row: 1, col: 2 });
	});
});

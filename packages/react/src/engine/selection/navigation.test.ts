import { describe, expect, it } from 'vitest';

import { cellKey, createWorksheet } from 'xlsx-viewer-core';

import { currentRegion, DataProbe, jumpToEdge, moveOne } from './navigation';
import { cycleActiveWithinSelection, singleCellSelection } from './selection';

function probeWith(cells: Array<[number, number]>, hiddenRows: number[] = []) {
	const sheet = createWorksheet('S', 1);
	for (const [r, c] of cells) {
		sheet.cells.set(cellKey(r, c), { value: 1 });
	}
	return new DataProbe(
		sheet,
		(r) => hiddenRows.includes(r),
		() => false,
	);
}

describe('jumpToEdge (Ctrl+Arrow)', () => {
	// Column A: rows 0-3 filled, gap, rows 7-8 filled.
	const probe = probeWith([
		[0, 0],
		[1, 0],
		[2, 0],
		[3, 0],
		[7, 0],
		[8, 0],
	]);

	it('runs to the end of a block', () => {
		expect(jumpToEdge(probe, { row: 0, col: 0 }, 'down')).toEqual({ row: 3, col: 0 });
	});

	it('jumps across a gap to the next block', () => {
		expect(jumpToEdge(probe, { row: 3, col: 0 }, 'down')).toEqual({ row: 7, col: 0 });
		expect(jumpToEdge(probe, { row: 5, col: 0 }, 'up')).toEqual({ row: 3, col: 0 });
	});

	it('goes to the sheet edge past the last data', () => {
		expect(jumpToEdge(probe, { row: 8, col: 0 }, 'down')).toEqual({ row: 1048575, col: 0 });
		expect(jumpToEdge(probe, { row: 0, col: 3 }, 'right')).toEqual({ row: 0, col: 16383 });
		expect(jumpToEdge(probe, { row: 0, col: 3 }, 'left')).toEqual({ row: 0, col: 0 });
	});
});

describe('moveOne', () => {
	it('skips hidden rows', () => {
		const probe = probeWith([], [1, 2]);
		expect(moveOne(probe, { row: 0, col: 0 }, 'down')).toEqual({ row: 3, col: 0 });
		expect(moveOne(probe, { row: 0, col: 0 }, 'up')).toEqual({ row: 0, col: 0 });
	});
});

describe('currentRegion', () => {
	it('expands over the connected block', () => {
		const probe = probeWith([
			[1, 1],
			[1, 2],
			[2, 1],
			[3, 3],
			[6, 6],
		]);
		expect(currentRegion(probe, { row: 1, col: 1 })).toEqual({
			top: 1,
			left: 1,
			bottom: 3,
			right: 3,
		});
	});
});

describe('cycleActiveWithinSelection', () => {
	it('wraps down then right inside the range', () => {
		const sel = {
			...singleCellSelection(0, 0),
			ranges: [{ top: 0, left: 0, bottom: 1, right: 1 }],
		};
		const a = cycleActiveWithinSelection(sel, 'down');
		expect(a.active).toEqual({ row: 1, col: 0 });
		const b = cycleActiveWithinSelection(a, 'down');
		expect(b.active).toEqual({ row: 0, col: 1 });
	});
});

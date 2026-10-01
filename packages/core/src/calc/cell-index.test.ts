import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import type { Cell } from '../types/cell';
import { lowerBound, SheetCellIndex } from './cell-index';

function cells(entries: Array<[number, number]>): Map<number, Cell> {
	return new Map(entries.map(([r, c]) => [cellKey(r, c), { value: r * 100 + c }]));
}

describe('SheetCellIndex', () => {
	it('indexes stored cells by column and tracks the used range', () => {
		const map = cells([
			[5, 2],
			[1, 2],
			[3, 0],
			[9, 7],
		]);
		const idx = new SheetCellIndex();
		idx.sync(map);
		expect(idx.count).toBe(4);
		expect(idx.rows(2)).toEqual([1, 5]);
		expect(idx.columnList()).toEqual([0, 2, 7]);
		expect(idx.usedRange()).toEqual({ top: 1, left: 0, bottom: 9, right: 7 });
		expect(idx.keysIn({ top: 0, left: 0, bottom: 6, right: 2 })).toEqual([
			cellKey(1, 2),
			cellKey(3, 0),
			cellKey(5, 2),
		]);
		expect(idx.countIn({ top: 0, left: 0, bottom: 6, right: 2 })).toBe(3);
		expect(idx.countIn({ top: 10, left: 0, bottom: 20, right: 9 })).toBe(0);
	});

	it('updates incrementally and shrinks the used range on deletion', () => {
		const map = cells([
			[0, 0],
			[4, 4],
		]);
		const idx = new SheetCellIndex();
		idx.sync(map);
		map.set(cellKey(8, 1), { value: 1 });
		idx.note(map, 8, 1);
		expect(idx.usedRange()).toEqual({ top: 0, left: 0, bottom: 8, right: 4 });
		map.delete(cellKey(8, 1));
		idx.note(map, 8, 1);
		map.delete(cellKey(4, 4));
		idx.note(map, 4, 4);
		expect(idx.usedRange()).toEqual({ top: 0, left: 0, bottom: 0, right: 0 });
		map.delete(cellKey(0, 0));
		idx.note(map, 0, 0);
		expect(idx.usedRange()).toBeNull();
	});

	it('rebuilds when the map is replaced or changed without notice', () => {
		const idx = new SheetCellIndex();
		const a = cells([[0, 0]]);
		idx.sync(a);
		const b = cells([
			[2, 3],
			[7, 3],
		]);
		idx.sync(b);
		expect(idx.rows(3)).toEqual([2, 7]);
		b.set(cellKey(9, 9), { value: 1 });
		idx.sync(b);
		expect(idx.usedRange()).toEqual({ top: 2, left: 3, bottom: 9, right: 9 });
	});

	it('lowerBound finds insertion points', () => {
		expect(lowerBound([1, 3, 5], 0)).toBe(0);
		expect(lowerBound([1, 3, 5], 3)).toBe(1);
		expect(lowerBound([1, 3, 5], 4)).toBe(2);
		expect(lowerBound([1, 3, 5], 9)).toBe(3);
	});
});

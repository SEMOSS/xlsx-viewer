import { describe, expect, it } from 'vitest';

import { FormulaColumn } from './formula-column';

describe('FormulaColumn', () => {
	it('enumerates dirty rows in a range like a brute-force scan', () => {
		const column = new FormulaColumn();
		const dirty = new Set<number>();
		for (let r = 0; r < 500; r += 3) {
			column.push(r);
		}
		column.sort();
		let seed = 7;
		const next = () => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31;
		for (let step = 0; step < 400; step++) {
			const row = column.rows[Math.floor(next() * column.rows.length)];
			if (dirty.has(row)) {
				dirty.delete(row);
				column.mark(row, -1);
			} else {
				dirty.add(row);
				column.mark(row, 1);
			}
			const top = Math.floor(next() * 500);
			const bottom = top + Math.floor(next() * 300);
			const got: number[] = [];
			column.forEachDirty(
				top,
				bottom,
				(r) => dirty.has(r),
				(r) => got.push(r),
			);
			const want = column.rows.filter((r) => r >= top && r <= bottom && dirty.has(r));
			expect(got).toEqual(want);
			expect(column.dirty).toBe(dirty.size);
		}
	});

	it('keeps counts across inserts and removals', () => {
		const column = new FormulaColumn();
		const dirty = new Set<number>([5, 50]);
		column.insert(50);
		column.insert(5);
		column.mark(5, 1);
		column.mark(50, 1);
		for (let r = 100; r < 200; r++) {
			column.insert(r);
		}
		const seen: number[] = [];
		column.forEachDirty(
			0,
			1000,
			(r) => dirty.has(r),
			(r) => seen.push(r),
		);
		expect(seen).toEqual([5, 50]);
		column.remove(5);
		dirty.delete(5);
		column.mark(5, -1);
		const after: number[] = [];
		column.forEachDirty(
			0,
			1000,
			(r) => dirty.has(r),
			(r) => after.push(r),
		);
		expect(after).toEqual([50]);
	});
});

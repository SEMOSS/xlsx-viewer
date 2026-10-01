import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import type { RangeAddress } from '../types/address';
import { RangeIndex } from './range-index';

/** Deterministic pseudo-random numbers (mulberry32). */
function rng(seed: number): () => number {
	let a = seed;
	return () => {
		a |= 0;
		a = (a + 0x6d2b79f5) | 0;
		let t = Math.imul(a ^ (a >>> 15), 1 | a);
		t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

function contains(r: RangeAddress, row: number, col: number): boolean {
	return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

function intersects(a: RangeAddress, b: RangeAddress): boolean {
	return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

function randomRange(next: () => number): RangeAddress {
	const kind = next();
	if (kind < 0.1) {
		const c = Math.floor(next() * 40);
		return { top: 0, left: c, bottom: MAX_ROW_INDEX, right: c + Math.floor(next() * 3) };
	}
	if (kind < 0.15) {
		const r = Math.floor(next() * 3000);
		return { top: r, left: 0, bottom: r, right: MAX_COL_INDEX };
	}
	const top = Math.floor(next() * 3000);
	const left = Math.floor(next() * 40);
	return {
		top,
		left,
		bottom: top + Math.floor(next() * (next() < 0.2 ? 5000 : 30)),
		right: left + Math.floor(next() * (next() < 0.1 ? 300 : 4)),
	};
}

describe('RangeIndex', () => {
	it('answers point and range queries exactly (against brute force)', () => {
		const next = rng(42);
		const index = new RangeIndex<number>();
		const ranges: RangeAddress[] = [];
		const entries = [];
		for (let i = 0; i < 1500; i++) {
			const r = randomRange(next);
			ranges.push(r);
			entries.push(index.insert(r, i));
		}
		// Remove a third of them.
		const removed = new Set<number>();
		for (let i = 0; i < 1500; i += 3) {
			index.remove(entries[i]);
			removed.add(i);
		}
		expect(index.size).toBe(1000);
		for (let q = 0; q < 300; q++) {
			const row = Math.floor(next() * 6000);
			const col = Math.floor(next() * 60);
			const got: number[] = [];
			index.queryPoint(row, col, (i) => got.push(i));
			const want = ranges
				.map((_, i) => i)
				.filter((i) => !removed.has(i) && contains(ranges[i], row, col));
			expect(got.sort((a, b) => a - b)).toEqual(want);
		}
		for (let q = 0; q < 200; q++) {
			const query = randomRange(next);
			const got: number[] = [];
			index.queryRange(query, (i) => got.push(i));
			const want = ranges
				.map((_, i) => i)
				.filter((i) => !removed.has(i) && intersects(ranges[i], query));
			expect(got.sort((a, b) => a - b)).toEqual(want);
		}
	});

	it('handles grid-sized rectangles and clearing', () => {
		const index = new RangeIndex<string>();
		index.insert({ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: MAX_COL_INDEX }, 'all');
		index.insert({ top: 1, left: 1, bottom: MAX_ROW_INDEX - 1, right: MAX_COL_INDEX - 1 }, 'inner');
		const hits: string[] = [];
		index.queryPoint(MAX_ROW_INDEX, MAX_COL_INDEX, (s) => hits.push(s));
		expect(hits).toEqual(['all']);
		index.clear();
		index.queryPoint(5, 5, (s) => hits.push(s));
		expect(hits).toEqual(['all']);
		expect(index.size).toBe(0);
	});

	it('removes entries idempotently', () => {
		const index = new RangeIndex<number>();
		const e = index.insert({ top: 0, left: 0, bottom: 9, right: 0 }, 1);
		index.remove(e);
		index.remove(e);
		let n = 0;
		index.queryPoint(3, 0, () => n++);
		expect(n).toBe(0);
	});
});

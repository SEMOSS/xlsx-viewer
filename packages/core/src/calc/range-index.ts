import type { RangeAddress } from '../types/address';

/**
 * Spatial index of rectangles answering "which entries contain this cell"
 * and "which entries intersect this range" without scanning every entry.
 *
 * Each rectangle is decomposed into canonical dyadic blocks (as in a 2-D
 * segment tree): rows into aligned power-of-two runs, columns likewise, and
 * the entry is registered in every (column block, row block) bucket. A
 * point lies in exactly one block per level pair, so a point query probes
 * one bucket per level pair in use and every entry found really contains
 * the point. Rectangles that would need too many buckets go to a short
 * list that is checked linearly.
 */

const ROW_LEVELS = 21; // 2^20 rows
const COL_LEVELS = 15; // 2^14 columns
const LEVEL_PAIRS = ROW_LEVELS * COL_LEVELS;
const BLOCK_SPACE = 2 ** 34;
const ROW_SPACE = 2 ** 20;
/** Registrations beyond this go to the linear list. */
const MAX_BUCKETS = 64;

export interface RangeIndexEntry<T> {
	readonly range: RangeAddress;
	readonly item: T;
	/** Bucket keys the entry is registered in; `null` for the linear list. */
	buckets: number[] | null;
	stamp: number;
	/** Last {@link RangeIndex.queryPoint} round that visited the entry. */
	round: number;
}

/** Canonical decomposition of [lo, hi] into aligned blocks: pairs (level, index). */
function decompose(lo: number, hi: number, maxLevel: number, out: number[]): void {
	out.length = 0;
	let p = lo;
	while (p <= hi) {
		let level = 0;
		while (level < maxLevel) {
			const size = 2 ** (level + 1);
			if (p % size !== 0 || p + size - 1 > hi) {
				break;
			}
			level++;
		}
		out.push(level, Math.floor(p / 2 ** level));
		p += 2 ** level;
	}
}

const rowBlocks: number[] = [];
const colBlocks: number[] = [];

export class RangeIndex<T> {
	private readonly buckets = new Map<number, Array<RangeIndexEntry<T>>>();
	private readonly levelCounts = new Int32Array(LEVEL_PAIRS);
	private active: number[] | null = [];
	private readonly linear: Array<RangeIndexEntry<T>> = [];
	private readonly all = new Set<RangeIndexEntry<T>>();
	private stamp = 0;
	/** Round in which each bucket was last visited by a round query. */
	private readonly visited = new Map<number, number>();

	get size(): number {
		return this.all.size;
	}

	insert(range: RangeAddress, item: T): RangeIndexEntry<T> {
		const entry: RangeIndexEntry<T> = { range, item, buckets: null, stamp: 0, round: 0 };
		this.all.add(entry);
		decompose(range.left, range.right, COL_LEVELS - 1, colBlocks);
		decompose(range.top, range.bottom, ROW_LEVELS - 1, rowBlocks);
		const count = (colBlocks.length / 2) * (rowBlocks.length / 2);
		if (count > MAX_BUCKETS) {
			this.linear.push(entry);
			return entry;
		}
		const keys: number[] = [];
		for (let i = 0; i < colBlocks.length; i += 2) {
			for (let j = 0; j < rowBlocks.length; j += 2) {
				const pair = colBlocks[i] * ROW_LEVELS + rowBlocks[j];
				const key = pair * BLOCK_SPACE + colBlocks[i + 1] * ROW_SPACE + rowBlocks[j + 1];
				let bucket = this.buckets.get(key);
				if (bucket === undefined) {
					bucket = [];
					this.buckets.set(key, bucket);
				}
				bucket.push(entry);
				if (this.levelCounts[pair]++ === 0) {
					this.active = null;
				}
				keys.push(key);
			}
		}
		entry.buckets = keys;
		return entry;
	}

	remove(entry: RangeIndexEntry<T>): void {
		if (!this.all.delete(entry)) {
			return;
		}
		if (entry.buckets === null) {
			const i = this.linear.indexOf(entry);
			if (i >= 0) {
				this.linear[i] = this.linear[this.linear.length - 1];
				this.linear.pop();
			}
			return;
		}
		for (const key of entry.buckets) {
			const bucket = this.buckets.get(key);
			if (bucket === undefined) {
				continue;
			}
			const i = bucket.indexOf(entry);
			if (i >= 0) {
				bucket[i] = bucket[bucket.length - 1];
				bucket.pop();
			}
			if (bucket.length === 0) {
				this.buckets.delete(key);
			}
			const pair = Math.floor(key / BLOCK_SPACE);
			if (--this.levelCounts[pair] === 0) {
				this.active = null;
			}
		}
		entry.buckets = null;
	}

	clear(): void {
		this.buckets.clear();
		this.visited.clear();
		this.levelCounts.fill(0);
		this.active = [];
		this.linear.length = 0;
		this.all.clear();
	}

	private activePairs(): number[] {
		if (this.active === null) {
			const list: number[] = [];
			for (let p = 0; p < LEVEL_PAIRS; p++) {
				if (this.levelCounts[p] > 0) {
					list.push(p);
				}
			}
			this.active = list;
		}
		return this.active;
	}

	/**
	 * Visit every entry whose rectangle contains (row, col). With a `round`
	 * number, entries already visited in that round are skipped, a whole
	 * bucket at a time: repeated queries of one round (marking the dependents
	 * of many cells) then cost O(buckets) instead of O(entries) each.
	 */
	queryPoint(
		row: number,
		col: number,
		visit: (item: T, entry: RangeIndexEntry<T>) => void,
		round = 0,
	): void {
		if (this.all.size === 0) {
			return;
		}
		for (const pair of this.activePairs()) {
			const cl = Math.floor(pair / ROW_LEVELS);
			const rl = pair - cl * ROW_LEVELS;
			const key =
				pair * BLOCK_SPACE + Math.floor(col / 2 ** cl) * ROW_SPACE + Math.floor(row / 2 ** rl);
			const bucket = this.buckets.get(key);
			if (bucket === undefined) {
				continue;
			}
			if (round !== 0) {
				if (this.visited.get(key) === round) {
					continue;
				}
				this.visited.set(key, round);
			}
			for (let i = 0; i < bucket.length; i++) {
				const e = bucket[i];
				if (round === 0) {
					visit(e.item, e);
				} else if (e.round !== round) {
					e.round = round;
					visit(e.item, e);
				}
			}
		}
		for (const e of this.linear) {
			const r = e.range;
			if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
				if (round === 0) {
					visit(e.item, e);
				} else if (e.round !== round) {
					e.round = round;
					visit(e.item, e);
				}
			}
		}
	}

	/** Visit every entry whose rectangle intersects `range` (each once). */
	queryRange(range: RangeAddress, visit: (item: T, entry: RangeIndexEntry<T>) => void): void {
		if (this.all.size === 0) {
			return;
		}
		const pairs = this.activePairs();
		let probes = 0;
		for (const pair of pairs) {
			const cl = Math.floor(pair / ROW_LEVELS);
			const rl = pair - cl * ROW_LEVELS;
			const cols = Math.floor(range.right / 2 ** cl) - Math.floor(range.left / 2 ** cl) + 1;
			const rows = Math.floor(range.bottom / 2 ** rl) - Math.floor(range.top / 2 ** rl) + 1;
			probes += cols * rows;
			if (probes > this.all.size) {
				break;
			}
		}
		if (probes > this.all.size) {
			for (const e of this.all) {
				const r = e.range;
				if (
					r.left <= range.right &&
					range.left <= r.right &&
					r.top <= range.bottom &&
					range.top <= r.bottom
				) {
					visit(e.item, e);
				}
			}
			return;
		}
		const stamp = ++this.stamp;
		for (const pair of pairs) {
			const cl = Math.floor(pair / ROW_LEVELS);
			const rl = pair - cl * ROW_LEVELS;
			const c0 = Math.floor(range.left / 2 ** cl);
			const c1 = Math.floor(range.right / 2 ** cl);
			const r0 = Math.floor(range.top / 2 ** rl);
			const r1 = Math.floor(range.bottom / 2 ** rl);
			for (let c = c0; c <= c1; c++) {
				for (let r = r0; r <= r1; r++) {
					const bucket = this.buckets.get(pair * BLOCK_SPACE + c * ROW_SPACE + r);
					if (bucket === undefined) {
						continue;
					}
					for (const e of bucket) {
						if (e.stamp !== stamp) {
							e.stamp = stamp;
							visit(e.item, e);
						}
					}
				}
			}
		}
		for (const e of this.linear) {
			const r = e.range;
			if (
				r.left <= range.right &&
				range.left <= r.right &&
				r.top <= range.bottom &&
				range.top <= r.bottom
			) {
				visit(e.item, e);
			}
		}
	}
}

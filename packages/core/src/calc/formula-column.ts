import { lowerBound } from './cell-index';

/**
 * Rows of the formula cells of one column, with a count of the dirty ones
 * kept in a Fenwick tree so "dirty formulas in rows a..b" is answered in
 * O(log n + found) instead of scanning every formula in the range (running
 * totals such as `SUM($A$1:A5000)` filled down would otherwise be quadratic).
 */
export class FormulaColumn {
	/** Sorted rows. */
	readonly rows: number[] = [];
	/** Number of dirty formulas in the column. */
	dirty = 0;
	/** Fenwick tree over row indices (1-based); null when it must be rebuilt. */
	private tree: Int32Array | null = null;
	private top = 1;

	/** Append without keeping order (bulk build); call `sort` afterwards. */
	push(row: number): void {
		this.rows.push(row);
		this.tree = null;
	}

	sort(): void {
		this.rows.sort((a, b) => a - b);
		this.tree = null;
	}

	insert(row: number): void {
		this.rows.splice(lowerBound(this.rows, row), 0, row);
		this.tree = null;
	}

	remove(row: number): void {
		const i = lowerBound(this.rows, row);
		if (this.rows[i] === row) {
			this.rows.splice(i, 1);
			this.tree = null;
		}
	}

	/** A formula of the column changed dirtiness (`delta` is +1 or -1). */
	mark(row: number, delta: number): void {
		this.dirty += delta;
		const tree = this.tree;
		if (tree !== null) {
			const i = lowerBound(this.rows, row);
			if (this.rows[i] === row) {
				for (let j = i + 1; j < tree.length; j += j & -j) {
					tree[j] += delta;
				}
			}
		}
	}

	private build(isDirty: (row: number) => boolean): Int32Array {
		const n = this.rows.length;
		const tree = new Int32Array(n + 1);
		for (let i = 1; i <= n; i++) {
			if (isDirty(this.rows[i - 1])) {
				tree[i] += 1;
			}
			const parent = i + (i & -i);
			if (parent <= n) {
				tree[parent] += tree[i];
			}
		}
		let top = 1;
		while (top * 2 <= n) {
			top *= 2;
		}
		this.top = top;
		this.tree = tree;
		return tree;
	}

	private prefix(tree: Int32Array, count: number): number {
		let s = 0;
		for (let i = count; i > 0; i -= i & -i) {
			s += tree[i];
		}
		return s;
	}

	/** 0-based index of the k-th (1-based) dirty formula. */
	private kth(tree: Int32Array, k: number): number {
		let pos = 0;
		let rest = k;
		for (let step = this.top; step > 0; step >>= 1) {
			const next = pos + step;
			if (next < tree.length && tree[next] < rest) {
				pos = next;
				rest -= tree[next];
			}
		}
		return pos;
	}

	/**
	 * Visit the rows of dirty formulas between `top` and `bottom`.
	 * `isDirty` reads a row's state (used to build the tree lazily).
	 */
	forEachDirty(
		top: number,
		bottom: number,
		isDirty: (row: number) => boolean,
		visit: (row: number) => void,
	): void {
		if (this.dirty === 0) {
			return;
		}
		const lo = lowerBound(this.rows, top);
		const hi = lowerBound(this.rows, bottom + 1);
		if (hi - lo <= 32) {
			for (let i = lo; i < hi; i++) {
				if (isDirty(this.rows[i])) {
					visit(this.rows[i]);
				}
			}
			return;
		}
		const tree = this.tree ?? this.build(isDirty);
		const before = this.prefix(tree, lo);
		const count = this.prefix(tree, hi) - before;
		for (let k = 1; k <= count; k++) {
			visit(this.rows[this.kth(tree, before + k)]);
		}
	}
}

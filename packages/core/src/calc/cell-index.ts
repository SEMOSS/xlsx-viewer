import { cellKey, keyCol, keyRow } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';

/** Rows of one column holding a cell, with a lazily sorted copy. */
interface ColumnRows {
	set: Set<number>;
	sorted: number[] | null;
}

/** First index in a sorted array whose value is >= `v`. */
export function lowerBound(a: readonly number[], v: number): number {
	let lo = 0;
	let hi = a.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if (a[mid] < v) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	return lo;
}

/**
 * Column index of the cells stored on a sheet plus their bounding box, so
 * large and whole-column references visit only cells that exist. The index
 * is rebuilt when the sheet's cell map is replaced (undo restores a copy) or
 * its size drifts without notification; otherwise `note` keeps it current.
 */
export class SheetCellIndex {
	private source: Map<number, Cell> | null = null;
	private size = 0;
	private readonly columns = new Map<number, ColumnRows>();
	private sortedColumns: number[] | null = null;
	private bounds: RangeAddress | null = null;
	private boundsValid = true;

	/** Bring the index in line with `cells` if it was replaced or changed behind our back. */
	sync(cells: Map<number, Cell>): void {
		if (this.source !== cells || this.size !== cells.size) {
			this.build(cells);
		}
	}

	private build(cells: Map<number, Cell>): void {
		this.source = cells;
		this.size = cells.size;
		this.columns.clear();
		this.sortedColumns = null;
		let top = Infinity;
		let left = Infinity;
		let bottom = -1;
		let right = -1;
		for (const key of cells.keys()) {
			const row = keyRow(key);
			const col = keyCol(key);
			this.addRow(col, row);
			if (row < top) {
				top = row;
			}
			if (row > bottom) {
				bottom = row;
			}
			if (col < left) {
				left = col;
			}
			if (col > right) {
				right = col;
			}
		}
		this.bounds = bottom < 0 ? null : { top, left, bottom, right };
		this.boundsValid = true;
	}

	private addRow(col: number, row: number): void {
		let c = this.columns.get(col);
		if (c === undefined) {
			c = { set: new Set(), sorted: null };
			this.columns.set(col, c);
			this.sortedColumns = null;
		}
		if (!c.set.has(row)) {
			c.set.add(row);
			const sorted = c.sorted;
			if (sorted !== null) {
				// Keep the sorted copy current (usually an append) instead of re-sorting later.
				if (sorted.length === 0 || sorted[sorted.length - 1] < row) {
					sorted.push(row);
				} else {
					sorted.splice(lowerBound(sorted, row), 0, row);
				}
			}
		}
	}

	/** Record that the cell at (row, col) was added, replaced or deleted. */
	note(cells: Map<number, Cell>, row: number, col: number): void {
		if (this.source !== cells) {
			this.build(cells);
			return;
		}
		this.size = cells.size;
		const present = cells.has(cellKey(row, col));
		const c = this.columns.get(col);
		if (present) {
			this.addRow(col, row);
			const b = this.bounds;
			if (b === null) {
				if (this.boundsValid) {
					this.bounds = { top: row, left: col, bottom: row, right: col };
				}
			} else {
				b.top = Math.min(b.top, row);
				b.left = Math.min(b.left, col);
				b.bottom = Math.max(b.bottom, row);
				b.right = Math.max(b.right, col);
			}
		} else if (c !== undefined && c.set.delete(row)) {
			if (c.sorted !== null) {
				c.sorted.splice(lowerBound(c.sorted, row), 1);
			}
			if (c.set.size === 0) {
				this.columns.delete(col);
				this.sortedColumns = null;
			}
			const b = this.bounds;
			if (b !== null && (row === b.top || row === b.bottom || col === b.left || col === b.right)) {
				this.boundsValid = false;
			}
		}
	}

	/** Number of stored cells. */
	get count(): number {
		return this.size;
	}

	/** Bounding box of the stored cells, or null for an empty sheet. */
	usedRange(): RangeAddress | null {
		if (!this.boundsValid) {
			let top = Infinity;
			let left = Infinity;
			let bottom = -1;
			let right = -1;
			for (const [col, c] of this.columns) {
				const rows = this.rowsOf(c);
				if (rows.length === 0) {
					continue;
				}
				top = Math.min(top, rows[0]);
				bottom = Math.max(bottom, rows[rows.length - 1]);
				left = Math.min(left, col);
				right = Math.max(right, col);
			}
			this.bounds = bottom < 0 ? null : { top, left, bottom, right };
			this.boundsValid = true;
		}
		return this.bounds === null ? null : { ...this.bounds };
	}

	private rowsOf(c: ColumnRows): number[] {
		if (c.sorted === null) {
			c.sorted = Array.from(c.set).sort((a, b) => a - b);
		}
		return c.sorted;
	}

	/** Sorted rows of the cells stored in a column. */
	rows(col: number): readonly number[] {
		const c = this.columns.get(col);
		return c === undefined ? [] : this.rowsOf(c);
	}

	/** Columns holding at least one cell, ascending. */
	columnList(): readonly number[] {
		if (this.sortedColumns === null) {
			this.sortedColumns = Array.from(this.columns.keys()).sort((a, b) => a - b);
		}
		return this.sortedColumns;
	}

	/** Number of stored cells inside `range` (binary searches per column). */
	countIn(range: RangeAddress): number {
		const cols = this.columnList();
		let n = 0;
		for (let i = lowerBound(cols, range.left); i < cols.length && cols[i] <= range.right; i++) {
			const rows = this.rows(cols[i]);
			n += lowerBound(rows, range.bottom + 1) - lowerBound(rows, range.top);
		}
		return n;
	}

	/**
	 * Keys of the stored cells inside `range`, in row-major order, found
	 * through the column index (for ranges much larger than their content).
	 */
	keysIn(range: RangeAddress): number[] {
		const cols = this.columnList();
		const out: number[] = [];
		let start = lowerBound(cols, range.left);
		let multi = 0;
		for (; start < cols.length && cols[start] <= range.right; start++) {
			const col = cols[start];
			const rows = this.rows(col);
			let i = lowerBound(rows, range.top);
			if (i < rows.length && rows[i] <= range.bottom) {
				multi++;
			}
			for (; i < rows.length && rows[i] <= range.bottom; i++) {
				out.push(cellKey(rows[i], col));
			}
		}
		if (multi > 1) {
			out.sort((a, b) => a - b);
		}
		return out;
	}
}

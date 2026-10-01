import { cellKey } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_COLS, MAX_ROW_INDEX } from '../address/constants';
import type { ArrayValue, RefValue, Scalar } from '../formula/types';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import { lowerBound, type SheetCellIndex } from './cell-index';

/** Ranges above this many cells are read through the column index rather than cell by cell. */
const DIRECT_SCAN_LIMIT = 1024;
/** Single-column reads at least this tall use (and build) a column snapshot during a pass. */
const SNAPSHOT_MIN_ROWS = 64;

/** Non-blank values of one column, by row (valid until a cell of the column is written). */
interface ColumnSnapshot {
	rows: number[];
	values: Scalar[];
}

export interface RangeReaderHost {
	index(sheet: number): SheetCellIndex;
	/** Suspend the evaluation when formulas in the range still have to be recalculated. */
	ensureClean(sheet: number, range: RangeAddress): void;
}

/**
 * Bulk reads of references: materialising arrays and visiting non-blank
 * cells. Large sparse ranges go through the column index; during a
 * recalculation pass tall single-column reads use per-column snapshots so
 * repeated aggregates over the same column (`COUNTIF(A:A, ...)` filled
 * down) do not look up every cell again.
 */
export class RangeReader {
	/** Snapshots are only trusted while a pass runs (all writes then go through the engine). */
	snapshots = false;
	private readonly columns = new Map<number, ColumnSnapshot>();
	/** Memoised data derived from references, by key; and the keys depending on each column. */
	private readonly memos = new Map<string, unknown>();
	private readonly memoColumns = new Map<number, string[]>();

	constructor(
		private readonly workbook: Workbook,
		private readonly host: RangeReaderHost,
	) {}

	/** Drop all snapshots and memos (pass boundaries, rebuilds). */
	clear(): void {
		this.columns.clear();
		this.memos.clear();
		this.memoColumns.clear();
	}

	/**
	 * The cell at (row, col) now holds `value` (null: blank or removed). The
	 * column snapshot is patched in place (formulas are usually written in
	 * row order, so this is an append); memos over the column are dropped.
	 */
	noteWrite(sheet: number, row: number, col: number, value: Scalar): void {
		const id = sheet * MAX_COLS + col;
		const snap = this.columns.get(id);
		if (snap !== undefined) {
			const i = lowerBound(snap.rows, row);
			const present = snap.rows[i] === row;
			if (value === null) {
				if (present) {
					snap.rows.splice(i, 1);
					snap.values.splice(i, 1);
				}
			} else if (present) {
				snap.values[i] = value;
			} else {
				snap.rows.splice(i, 0, row);
				snap.values.splice(i, 0, value);
			}
		}
		if (this.memoColumns.size > 0) {
			this.dropMemos(id);
			this.dropMemos(-(sheet + 1));
		}
	}

	private dropMemos(id: number): void {
		const keys = this.memoColumns.get(id);
		if (keys !== undefined) {
			for (const k of keys) {
				this.memos.delete(k);
			}
			this.memoColumns.delete(id);
		}
	}

	/**
	 * Data derived from the values of a reference (a lookup index, say),
	 * built once per recalculation pass and reused until a cell of the
	 * reference's columns is written. Outside passes nothing is cached.
	 */
	memo<T>(ref: RefValue, key: string, build: () => T): T {
		const r = this.clamp(ref.sheet, ref.range);
		this.host.ensureClean(ref.sheet, r);
		if (!this.snapshots || (ref.sheetEnd ?? ref.sheet) !== ref.sheet) {
			return build();
		}
		const id = `${ref.sheet}:${ref.range.top}:${ref.range.left}:${ref.range.bottom}:${ref.range.right}|${key}`;
		if (this.memos.has(id)) {
			return this.memos.get(id) as T;
		}
		const value = build();
		this.memos.set(id, value);
		// Whole rows depend on every column of the sheet (the used range can grow).
		const wide = ref.range.right === MAX_COL_INDEX || r.right - r.left > 64;
		const columns = wide ? [-(ref.sheet + 1)] : [];
		for (let c = r.left; !wide && c <= r.right; c++) {
			columns.push(ref.sheet * MAX_COLS + c);
		}
		for (const colId of columns) {
			let keys = this.memoColumns.get(colId);
			if (keys === undefined) {
				keys = [];
				this.memoColumns.set(colId, keys);
			}
			keys.push(id);
		}
		return value;
	}

	private snapshot(sheet: number, col: number): ColumnSnapshot {
		const id = sheet * MAX_COLS + col;
		let snap = this.columns.get(id);
		if (snap === undefined) {
			const cells = this.workbook.sheets[sheet].cells;
			const rows: number[] = [];
			const values: Scalar[] = [];
			for (const row of this.host.index(sheet).rows(col)) {
				const v = cells.get(cellKey(row, col))?.value;
				if (v !== undefined && v !== null) {
					rows.push(row);
					values.push(v);
				}
			}
			snap = { rows, values };
			this.columns.set(id, snap);
		}
		return snap;
	}

	/** Clamp references that run to the edge of the grid to the used range. */
	clamp(sheet: number, r: RangeAddress): RangeAddress {
		if (r.bottom !== MAX_ROW_INDEX && r.right !== MAX_COL_INDEX) {
			return r;
		}
		const used = this.host.index(sheet).usedRange();
		const bottom =
			r.bottom === MAX_ROW_INDEX
				? Math.max(r.top, Math.min(r.bottom, used?.bottom ?? r.top))
				: r.bottom;
		const right =
			r.right === MAX_COL_INDEX
				? Math.max(r.left, Math.min(r.right, used?.right ?? r.left))
				: r.right;
		return { top: r.top, left: r.left, bottom, right };
	}

	private useSnapshot(r: RangeAddress): boolean {
		return this.snapshots && r.left === r.right && r.bottom - r.top + 1 >= SNAPSHOT_MIN_ROWS;
	}

	/** The range holds few stored cells for its size: iterate them through the column index. */
	private sparse(idx: SheetCellIndex, r: RangeAddress): boolean {
		const area = (r.bottom - r.top + 1) * (r.right - r.left + 1);
		return area > DIRECT_SCAN_LIMIT && idx.countIn(r) * 4 < area;
	}

	resolveRef(ref: RefValue): ArrayValue {
		const ws = this.workbook.sheets[ref.sheet];
		if (ws === undefined) {
			return { kind: 'array', rows: 1, cols: 1, values: [[ERRORS.REF]] };
		}
		const r = this.clamp(ref.sheet, ref.range);
		this.host.ensureClean(ref.sheet, r);
		const rows = r.bottom - r.top + 1;
		const cols = r.right - r.left + 1;
		const values: Scalar[][] = new Array(rows);
		const cells = ws.cells;
		const idx = this.host.index(ref.sheet);
		if (this.useSnapshot(r) || this.sparse(idx, r)) {
			for (let i = 0; i < rows; i++) {
				values[i] = new Array<Scalar>(cols).fill(null);
			}
			if (this.useSnapshot(r)) {
				const snap = this.snapshot(ref.sheet, r.left);
				for (
					let i = lowerBound(snap.rows, r.top);
					i < snap.rows.length && snap.rows[i] <= r.bottom;
					i++
				) {
					values[snap.rows[i] - r.top][0] = snap.values[i];
				}
			} else {
				for (const key of idx.keysIn(r)) {
					const cell = cells.get(key);
					if (cell !== undefined) {
						const row = Math.floor(key / MAX_COLS);
						values[row - r.top][key - row * MAX_COLS - r.left] = cell.value;
					}
				}
			}
			return { kind: 'array', rows, cols, values };
		}
		for (let i = 0; i < rows; i++) {
			const line: Scalar[] = new Array(cols);
			const base = cellKey(r.top + i, r.left);
			for (let j = 0; j < cols; j++) {
				line[j] = cells.get(base + j)?.value ?? null;
			}
			values[i] = line;
		}
		return { kind: 'array', rows, cols, values };
	}

	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void {
		const last = ref.sheetEnd ?? ref.sheet;
		for (let s = ref.sheet; s <= last; s++) {
			const ws = this.workbook.sheets[s];
			if (ws === undefined) {
				continue;
			}
			const idx = this.host.index(s);
			const used = idx.usedRange();
			if (used === null) {
				continue;
			}
			const top = Math.max(ref.range.top, used.top);
			const left = Math.max(ref.range.left, used.left);
			const bottom = Math.min(ref.range.bottom, used.bottom);
			const right = Math.min(ref.range.right, used.right);
			if (top > bottom || left > right) {
				continue;
			}
			const r = { top, left, bottom, right };
			this.host.ensureClean(s, r);
			if (this.useSnapshot(r)) {
				const snap = this.snapshot(s, left);
				for (
					let i = lowerBound(snap.rows, top);
					i < snap.rows.length && snap.rows[i] <= bottom;
					i++
				) {
					visit(snap.values[i], snap.rows[i], left, s);
				}
				continue;
			}
			const cells = ws.cells;
			if (!this.sparse(idx, r)) {
				for (let row = top; row <= bottom; row++) {
					const base = cellKey(row, 0);
					for (let col = left; col <= right; col++) {
						const v = cells.get(base + col)?.value;
						if (v !== undefined && v !== null) {
							visit(v, row, col, s);
						}
					}
				}
				continue;
			}
			for (const key of idx.keysIn(r)) {
				const v = cells.get(key)?.value;
				if (v !== undefined && v !== null) {
					const row = Math.floor(key / MAX_COLS);
					visit(v, row, key - row * MAX_COLS, s);
				}
			}
		}
	}
}

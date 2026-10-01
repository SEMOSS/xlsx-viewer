import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import type { RangeAddress } from '../../types/address';
import type { RefArea } from './tokens';

/** Rows or columns inserted / deleted on one sheet. `at` is 0-based. */
export interface StructureChange {
	kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';
	/** Sheet whose rows / columns change. */
	sheet: string;
	/** First inserted / deleted row or column (0-based). */
	at: number;
	count: number;
}

/** Result of rewriting a formula. */
export interface FormulaAdjustment {
	formula: string;
	/** `formula` differs from the input. */
	changed: boolean;
}

function affectsRows(change: StructureChange): boolean {
	return change.kind === 'insertRows' || change.kind === 'deleteRows';
}

function isInsert(change: StructureChange): boolean {
	return change.kind === 'insertRows' || change.kind === 'insertCols';
}

/**
 * Inserting `count` lines at `at` into the span `[lo, hi]`: a span starting
 * at or after `at` shifts, a span with `lo < at <= hi` grows, anything
 * before is untouched. Spans pushed past the grid lose their tail; `null`
 * when the whole span leaves the grid.
 */
function insertSpan(
	lo: number,
	hi: number,
	at: number,
	count: number,
	max: number,
): [number, number] | null {
	if (at <= lo) {
		lo += count;
		hi += count;
	} else if (at <= hi) {
		hi += count;
	} else {
		return [lo, hi];
	}
	return lo > max ? null : [lo, Math.min(hi, max)];
}

/**
 * Deleting lines `[at, at + count)` from the span `[lo, hi]`: spans inside
 * the band disappear (`null`), spans overlapping it shrink, spans after it
 * shift back.
 */
function deleteSpan(lo: number, hi: number, at: number, count: number): [number, number] | null {
	const last = at + count - 1;
	if (hi < at) {
		return [lo, hi];
	}
	if (lo > last) {
		return [lo - count, hi - count];
	}
	if (lo >= at && hi <= last) {
		return null;
	}
	return [lo < at ? lo : at, hi > last ? hi - count : at - 1];
}

function adjustSpan(lo: number, hi: number, change: StructureChange): [number, number] | null {
	if (isInsert(change)) {
		return insertSpan(
			lo,
			hi,
			change.at,
			change.count,
			affectsRows(change) ? MAX_ROW_INDEX : MAX_COL_INDEX,
		);
	}
	return deleteSpan(lo, hi, change.at, change.count);
}

/**
 * Adjust a range (merged cells, conditional formats, validations, tables,
 * names) for a row / column insertion or deletion on its sheet, following
 * Excel: insertion inside a range expands it, deletion shrinks it, a range
 * entirely deleted returns `null`. Whole-column ranges are unaffected by
 * row changes and whole-row ranges by column changes. Returns the input
 * object when nothing changes.
 */
export function adjustRangeForStructureChange(
	range: RangeAddress,
	change: StructureChange,
): RangeAddress | null {
	if (change.count <= 0) {
		return range;
	}
	if (affectsRows(change)) {
		if (range.top === 0 && range.bottom === MAX_ROW_INDEX) {
			return range;
		}
		const span = adjustSpan(range.top, range.bottom, change);
		if (span === null) {
			return null;
		}
		return span[0] === range.top && span[1] === range.bottom
			? range
			: { ...range, top: span[0], bottom: span[1] };
	}
	if (range.left === 0 && range.right === MAX_COL_INDEX) {
		return range;
	}
	const span = adjustSpan(range.left, range.right, change);
	if (span === null) {
		return null;
	}
	return span[0] === range.left && span[1] === range.right
		? range
		: { ...range, left: span[0], right: span[1] };
}

/**
 * Same rules for a reference area as written (corner order and `$` flags
 * are kept). Returns the input area when unchanged, `null` when deleted.
 */
export function adjustAreaForStructureChange(
	area: RefArea,
	change: StructureChange,
): RefArea | null {
	const rows = affectsRows(change);
	if (change.count <= 0 || (rows && area.kind === 'columns') || (!rows && area.kind === 'rows')) {
		return area;
	}
	const a = rows ? area.start.row : area.start.col;
	const b = rows ? area.end.row : area.end.col;
	const lo = Math.min(a, b);
	const hi = Math.max(a, b);
	if (lo === 0 && hi === (rows ? MAX_ROW_INDEX : MAX_COL_INDEX)) {
		// `A1:B1048576` spans every row, like `A:B`.
		return area;
	}
	const span = adjustSpan(lo, hi, change);
	if (span === null) {
		return null;
	}
	if (span[0] === lo && span[1] === hi) {
		return area;
	}
	const na = a <= b ? span[0] : span[1];
	const nb = a <= b ? span[1] : span[0];
	const start = rows ? { ...area.start, row: na } : { ...area.start, col: na };
	if (area.kind === 'cell') {
		return { kind: 'cell', start, end: start };
	}
	const end = rows ? { ...area.end, row: nb } : { ...area.end, col: nb };
	return { kind: area.kind, start, end };
}

import type { CellAddress, RangeAddress } from '../types/address';
import { columnToLetters, lettersToColumn, parseCellAddress } from './cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX, MAX_ROWS } from './constants';

/** Build a normalised range from two corners. */
export function makeRange(a: CellAddress, b: CellAddress = a): RangeAddress {
	return {
		top: Math.min(a.row, b.row),
		left: Math.min(a.col, b.col),
		bottom: Math.max(a.row, b.row),
		right: Math.max(a.col, b.col),
	};
}

/** Single-cell range. */
export function cellRange(row: number, col: number): RangeAddress {
	return { top: row, left: col, bottom: row, right: col };
}

export function rangeRowCount(r: RangeAddress): number {
	return r.bottom - r.top + 1;
}

export function rangeColCount(r: RangeAddress): number {
	return r.right - r.left + 1;
}

export function rangeCellCount(r: RangeAddress): number {
	return rangeRowCount(r) * rangeColCount(r);
}

export function isSingleCell(r: RangeAddress): boolean {
	return r.top === r.bottom && r.left === r.right;
}

/** The range spans every row (a whole-column reference such as `A:C`). */
export function isWholeColumnRange(r: RangeAddress): boolean {
	return r.top === 0 && r.bottom === MAX_ROW_INDEX;
}

/** The range spans every column (a whole-row reference such as `2:5`). */
export function isWholeRowRange(r: RangeAddress): boolean {
	return r.left === 0 && r.right === MAX_COL_INDEX;
}

export function rangeContains(r: RangeAddress, row: number, col: number): boolean {
	return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

export function rangeContainsRange(outer: RangeAddress, inner: RangeAddress): boolean {
	return (
		inner.top >= outer.top &&
		inner.bottom <= outer.bottom &&
		inner.left >= outer.left &&
		inner.right <= outer.right
	);
}

export function rangesIntersect(a: RangeAddress, b: RangeAddress): boolean {
	return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** Intersection of two ranges, or `null` when they do not overlap. */
export function rangeIntersection(a: RangeAddress, b: RangeAddress): RangeAddress | null {
	if (!rangesIntersect(a, b)) {
		return null;
	}
	return {
		top: Math.max(a.top, b.top),
		left: Math.max(a.left, b.left),
		bottom: Math.min(a.bottom, b.bottom),
		right: Math.min(a.right, b.right),
	};
}

/** Smallest range covering both inputs. */
export function rangeUnion(a: RangeAddress, b: RangeAddress): RangeAddress {
	return {
		top: Math.min(a.top, b.top),
		left: Math.min(a.left, b.left),
		bottom: Math.max(a.bottom, b.bottom),
		right: Math.max(a.right, b.right),
	};
}

export function rangesEqual(a: RangeAddress, b: RangeAddress): boolean {
	return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

/**
 * Format a range in A1 notation. Single cells print as `A1`, whole columns as
 * `A:C` and whole rows as `2:5` unless `forceCells` is set.
 */
export function formatRange(
	r: RangeAddress,
	opts: { absolute?: boolean; forceCells?: boolean } = {},
): string {
	const d = opts.absolute ? '$' : '';
	if (!opts.forceCells) {
		if (isWholeColumnRange(r) && !isWholeRowRange(r)) {
			return `${d}${columnToLetters(r.left)}:${d}${columnToLetters(r.right)}`;
		}
		if (isWholeRowRange(r) && !isWholeColumnRange(r)) {
			return `${d}${r.top + 1}:${d}${r.bottom + 1}`;
		}
	}
	const start = `${d}${columnToLetters(r.left)}${d}${r.top + 1}`;
	if (isSingleCell(r)) {
		return start;
	}
	return `${start}:${d}${columnToLetters(r.right)}${d}${r.bottom + 1}`;
}

/**
 * Parse `A1`, `A1:B2`, `$A$1:$B$2`, `A:C`, `2:5`. Corners are normalised.
 * Sheet prefixes are not accepted here (see the formula tokenizer for those).
 */
export function parseRange(ref: string): RangeAddress | null {
	const text = ref.trim();
	const colon = text.indexOf(':');
	if (colon < 0) {
		const cell = parseCellAddress(text);
		return cell ? cellRange(cell.row, cell.col) : null;
	}
	const a = text.slice(0, colon).replace(/\$/g, '');
	const b = text.slice(colon + 1).replace(/\$/g, '');
	if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
		const c1 = lettersToColumn(a);
		const c2 = lettersToColumn(b);
		if (c1 < 0 || c2 < 0) {
			return null;
		}
		return { top: 0, bottom: MAX_ROW_INDEX, left: Math.min(c1, c2), right: Math.max(c1, c2) };
	}
	if (/^\d{1,7}$/.test(a) && /^\d{1,7}$/.test(b)) {
		const r1 = Number(a) - 1;
		const r2 = Number(b) - 1;
		if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) {
			return null;
		}
		return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: 0, right: MAX_COL_INDEX };
	}
	const s = parseCellAddress(a);
	const e = parseCellAddress(b);
	if (!s || !e) {
		return null;
	}
	return makeRange(s, e);
}

/** Parse a space-separated list of ranges (`sqref`), e.g. `A1:B2 D4 F1:F9`. */
export function parseSqref(sqref: string): RangeAddress[] {
	const out: RangeAddress[] = [];
	for (const part of sqref.trim().split(/\s+/)) {
		if (!part) {
			continue;
		}
		const r = parseRange(part);
		if (r) {
			out.push(r);
		}
	}
	return out;
}

/** Format ranges as a space-separated `sqref`. */
export function formatSqref(ranges: readonly RangeAddress[]): string {
	return ranges.map((r) => formatRange(r, { forceCells: false })).join(' ');
}

/** Iterate every cell of a range in row-major order. */
export function* iterateRange(r: RangeAddress): Generator<CellAddress> {
	for (let row = r.top; row <= r.bottom; row++) {
		for (let col = r.left; col <= r.right; col++) {
			yield { row, col };
		}
	}
}

/** Clamp a range to the given used bounds (useful for whole-column references). */
export function clampRange(r: RangeAddress, bounds: RangeAddress): RangeAddress | null {
	return rangeIntersection(r, bounds);
}

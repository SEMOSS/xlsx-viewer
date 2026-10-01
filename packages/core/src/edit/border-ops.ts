/**
 * Format Cells > Border: set, remove or keep each edge of a range (outline,
 * inside lines and diagonals) with any line style and color, and read the
 * current edges of a range for the dialog preview.
 */
import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { getColInfo } from '../sheet/worksheet';
import type { RangeAddress } from '../types/address';
import type { Worksheet } from '../types/sheet';
import type { BorderEdge, BorderSpec } from '../types/style';
import type { Workbook } from '../types/workbook';
import { applyStylePatch } from './style-ops';
import type { EditTransaction } from './transaction';

/** The edges the Border tab controls. */
export type BorderEdgeName =
	| 'top'
	| 'bottom'
	| 'left'
	| 'right'
	| 'insideHorizontal'
	| 'insideVertical'
	| 'diagonalUp'
	| 'diagonalDown';

export const BORDER_EDGE_NAMES: readonly BorderEdgeName[] = [
	'top',
	'bottom',
	'left',
	'right',
	'insideHorizontal',
	'insideVertical',
	'diagonalUp',
	'diagonalDown',
];

/** Per edge: a line to draw, `null` to remove the line, absent to leave it unchanged. */
export type BorderEdgesChange = Partial<Record<BorderEdgeName, BorderEdge | null>>;

/** Current state of an edge across a range: one line, none, or different lines (`mixed`). */
export type RangeBorderState = Record<BorderEdgeName, BorderEdge | null | 'mixed'>;

const NONE: BorderEdge = { style: 'none' };

function asPatchEdge(edge: BorderEdge | null | undefined): BorderEdge | undefined {
	return edge === undefined ? undefined : edge === null ? NONE : edge;
}

function isWholeColumns(r: RangeAddress): boolean {
	return r.top === 0 && r.bottom === MAX_ROW_INDEX;
}

function isWholeRows(r: RangeAddress): boolean {
	return r.left === 0 && r.right === MAX_COL_INDEX;
}

interface Part {
	from: number;
	to: number;
	first: boolean;
	last: boolean;
}

/** First / middle / last slices of a span; a whole-sheet span is one inner slice. */
function parts(from: number, to: number, whole: boolean): Part[] {
	if (whole) {
		return [{ from, to, first: false, last: false }];
	}
	if (from === to) {
		return [{ from, to, first: true, last: true }];
	}
	const out: Part[] = [{ from, to: from, first: true, last: false }];
	if (to - from > 1) {
		out.push({ from: from + 1, to: to - 1, first: false, last: false });
	}
	out.push({ from: to, to, first: false, last: true });
	return out;
}

function diagonalPatch(change: BorderEdgesChange): Partial<BorderSpec> {
	const up = change.diagonalUp;
	const down = change.diagonalDown;
	const out: Partial<BorderSpec> = {};
	if (up === undefined && down === undefined) {
		return out;
	}
	if (up !== undefined) {
		out.diagonalUp = up !== null;
	}
	if (down !== undefined) {
		out.diagonalDown = down !== null;
	}
	const line = down ?? up;
	if (line) {
		out.diagonal = line;
	} else if (up === null && down === null) {
		out.diagonal = NONE;
	}
	return out;
}

function hasEdge(b: BorderSpec | undefined, side: 'top' | 'bottom' | 'left' | 'right'): boolean {
	const e = b?.[side];
	return e !== undefined && e.style !== 'none';
}

/** Existing cells of a range, iterating whichever is smaller (the range or the cell map). */
function existingKeys(sheet: Worksheet, range: RangeAddress): number[] {
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	const out: number[] = [];
	if (area > sheet.cells.size) {
		for (const key of sheet.cells.keys()) {
			const row = keyRow(key);
			const col = keyCol(key);
			if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) {
				out.push(key);
			}
		}
		return out;
	}
	for (let row = range.top; row <= range.bottom; row++) {
		for (let col = range.left; col <= range.right; col++) {
			const key = cellKey(row, col);
			if (sheet.cells.has(key)) {
				out.push(key);
			}
		}
	}
	return out;
}

/**
 * Cells just outside an outer edge that draw the same gridline: when that
 * edge is removed or restyled, theirs follows so the line looks the way the
 * user chose (Excel keeps both sides of a shared border consistent).
 */
function syncNeighbours(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	change: BorderEdgesChange,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const styles = workbook.styles;
	const borderOf = (key: number) =>
		styles.borders[styles.cellXfs[sheet.cells.get(key)?.style ?? 0]?.borderId ?? 0];
	const sides: Array<{
		edge: BorderEdge | null | undefined;
		strip: RangeAddress | null;
		facing: 'top' | 'bottom' | 'left' | 'right';
	}> = [
		{
			edge: change.top,
			facing: 'bottom',
			strip:
				!isWholeColumns(range) && range.top > 0
					? { ...range, top: range.top - 1, bottom: range.top - 1 }
					: null,
		},
		{
			edge: change.bottom,
			facing: 'top',
			strip:
				!isWholeColumns(range) && range.bottom < MAX_ROW_INDEX
					? { ...range, top: range.bottom + 1, bottom: range.bottom + 1 }
					: null,
		},
		{
			edge: change.left,
			facing: 'right',
			strip:
				!isWholeRows(range) && range.left > 0
					? { ...range, left: range.left - 1, right: range.left - 1 }
					: null,
		},
		{
			edge: change.right,
			facing: 'left',
			strip:
				!isWholeRows(range) && range.right < MAX_COL_INDEX
					? { ...range, left: range.right + 1, right: range.right + 1 }
					: null,
		},
	];
	for (const side of sides) {
		if (side.edge === undefined || !side.strip) {
			continue;
		}
		const targets = existingKeys(sheet, side.strip)
			.filter((key) => hasEdge(borderOf(key), side.facing))
			.map((key) => ({
				top: keyRow(key),
				bottom: keyRow(key),
				left: keyCol(key),
				right: keyCol(key),
			}));
		if (targets.length) {
			applyStylePatch(tx, workbook, sheetIndex, targets, {
				border: { [side.facing]: asPatchEdge(side.edge) },
			});
		}
	}
}

/**
 * Apply Border tab choices to a range: outer edges go to the cells on the
 * range's sides, inside lines to both cells sharing each inner gridline,
 * diagonals to every cell. Edges left out of `change` keep their current
 * lines. Whole rows / columns use row and column formats (their far edge
 * at the end of the sheet is not drawn, as nobody can see it).
 */
export function applyBorderEdges(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	change: BorderEdgesChange,
): void {
	if (BORDER_EDGE_NAMES.every((name) => change[name] === undefined)) {
		return;
	}
	const wholeCols = isWholeColumns(range);
	const wholeRows = isWholeRows(range);
	const diagonals = diagonalPatch(change);
	for (const rp of parts(range.top, range.bottom, wholeCols)) {
		for (const cp of parts(range.left, range.right, wholeRows)) {
			const patch: Partial<BorderSpec> = { ...diagonals };
			const top = asPatchEdge(rp.first ? change.top : change.insideHorizontal);
			const bottom = asPatchEdge(rp.last ? change.bottom : change.insideHorizontal);
			const left = asPatchEdge(cp.first ? change.left : change.insideVertical);
			const right = asPatchEdge(cp.last ? change.right : change.insideVertical);
			if (top) {
				patch.top = top;
			}
			if (bottom) {
				patch.bottom = bottom;
			}
			if (left) {
				patch.left = left;
			}
			if (right) {
				patch.right = right;
			}
			if (Object.keys(patch).length === 0) {
				continue;
			}
			const sub = { top: rp.from, bottom: rp.to, left: cp.from, right: cp.to };
			applyStylePatch(tx, workbook, sheetIndex, [sub], { border: patch });
		}
	}
	// The top row / left column of whole columns / rows still get their outer edge.
	if (wholeCols && change.top !== undefined) {
		applyStylePatch(tx, workbook, sheetIndex, [{ ...range, top: 0, bottom: 0 }], {
			border: { top: asPatchEdge(change.top) },
		});
	}
	if (wholeRows && change.left !== undefined) {
		applyStylePatch(tx, workbook, sheetIndex, [{ ...range, left: 0, right: 0 }], {
			border: { left: asPatchEdge(change.left) },
		});
	}
	syncNeighbours(tx, workbook, sheetIndex, range, change);
}

// ── Reading ──────────────────────────────────────────────────────────

/** Most lines (rows or columns) sampled per direction. */
const SAMPLE_LINES = 150;

function sameEdge(a: BorderEdge | null, b: BorderEdge | null): boolean {
	if (a === null || b === null) {
		return a === b;
	}
	return a.style === b.style && JSON.stringify(a.color ?? null) === JSON.stringify(b.color ?? null);
}

class EdgeAccumulator {
	private value: BorderEdge | null | 'mixed' | undefined;

	add(edge: BorderEdge | null): void {
		if (this.value === 'mixed') {
			return;
		}
		if (this.value === undefined) {
			this.value = edge;
		} else if (!sameEdge(this.value, edge)) {
			this.value = 'mixed';
		}
	}

	get result(): BorderEdge | null | 'mixed' {
		return this.value ?? null;
	}
}

/**
 * Lines of a span worth inspecting: both ends, the part holding cells and
 * the default-formatted lines next to it, thinned out evenly when long.
 */
function sampleLines(from: number, to: number, usedFrom: number, usedTo: number): number[] {
	const set = new Set<number>([from, to]);
	const lo = Math.max(from, usedFrom);
	const hi = Math.min(to, usedTo);
	const span = hi - lo + 1;
	const step = span > SAMPLE_LINES ? span / SAMPLE_LINES : 1;
	for (let i = 0; lo + Math.floor(i * step) <= hi && i < SAMPLE_LINES; i++) {
		set.add(lo + Math.floor(i * step));
	}
	if (hi >= lo) {
		set.add(hi);
	}
	if (lo - 1 >= from) {
		set.add(lo - 1);
	}
	if (hi + 1 <= to) {
		set.add(hi + 1);
	}
	return [...set].sort((a, b) => a - b);
}

function usedBounds(sheet: Worksheet): RangeAddress | null {
	if (sheet.cells.size === 0) {
		return null;
	}
	const b = { top: Infinity, left: Infinity, bottom: -1, right: -1 };
	for (const key of sheet.cells.keys()) {
		const row = keyRow(key);
		const col = keyCol(key);
		b.top = Math.min(b.top, row);
		b.bottom = Math.max(b.bottom, row);
		b.left = Math.min(b.left, col);
		b.right = Math.max(b.right, col);
	}
	return b;
}

/**
 * The edges of a range as the Border tab shows them: each outer side, the
 * inside lines and the diagonals are one line, none (`null`) or `mixed`
 * when the cells differ. Large ranges are sampled (both ends of every
 * side are always inspected).
 */
export function describeRangeBorders(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
): RangeBorderState {
	const sheet = workbook.sheets[sheetIndex];
	const styles = workbook.styles;
	const borderAt = (row: number, col: number): BorderSpec => {
		const cell = sheet.cells.get(cellKey(row, col));
		let xf = cell?.style;
		if (xf === undefined) {
			const info = sheet.rows.get(row);
			xf =
				info?.customFormat && info.style !== undefined ? info.style : getColInfo(sheet, col)?.style;
		}
		return styles.borders[styles.cellXfs[xf ?? 0]?.borderId ?? 0] ?? {};
	};
	const line = (e: BorderEdge | undefined): BorderEdge | null =>
		e && e.style !== 'none' ? e : null;
	const used = usedBounds(sheet) ?? range;
	const rows = sampleLines(range.top, range.bottom, used.top, used.bottom);
	const cols = sampleLines(range.left, range.right, used.left, used.right);
	const acc = Object.fromEntries(
		BORDER_EDGE_NAMES.map((n) => [n, new EdgeAccumulator()]),
	) as Record<BorderEdgeName, EdgeAccumulator>;
	for (const row of rows) {
		for (const col of cols) {
			const b = borderAt(row, col);
			if (row === range.top) {
				acc.top.add(line(b.top));
			}
			if (row === range.bottom) {
				acc.bottom.add(line(b.bottom));
			}
			if (col === range.left) {
				acc.left.add(line(b.left));
			}
			if (col === range.right) {
				acc.right.add(line(b.right));
			}
			if (row < range.bottom) {
				acc.insideHorizontal.add(line(b.bottom) ?? line(borderAt(row + 1, col).top));
			}
			if (col < range.right) {
				acc.insideVertical.add(line(b.right) ?? line(borderAt(row, col + 1).left));
			}
			acc.diagonalUp.add(b.diagonalUp ? line(b.diagonal) : null);
			acc.diagonalDown.add(b.diagonalDown ? line(b.diagonal) : null);
		}
	}
	const out = {} as RangeBorderState;
	for (const name of BORDER_EDGE_NAMES) {
		out[name] = acc[name].result;
	}
	return out;
}

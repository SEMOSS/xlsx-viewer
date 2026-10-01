import { keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { rangeContainsRange, rangesIntersect, rangeUnion } from '../address/range';
import { EMU_PER_POINT } from '../styles/column-width';
import type { RangeAddress } from '../types/address';
import type { DrawingAnchor } from '../types/drawing';
import type { Worksheet } from '../types/sheet';
import type { StyleSheet } from '../types/style';
import { printColWidth, printRowHeight, type AxisSize } from './metrics';

/**
 * What Excel prints when a sheet has no print area: every cell with a value,
 * a formula or visible formatting (a fill or a border), merged areas
 * around them, and the floating objects (pictures, charts, shapes).
 */

const MAX_WALK = 20_000;

/** Cells (and the rest of a partial row or column) covered by `extent` points from `start`. */
function walk(size: AxisSize, start: number, offset: number, extent: number, max: number): number {
	let pos = -offset;
	let i = start;
	for (let n = 0; n < MAX_WALK && i < max; n++, i++) {
		pos += size(i);
		if (pos >= extent) {
			return i;
		}
	}
	return Math.min(i, max);
}

/** Cells a drawing anchor covers. */
export function anchorCellRange(sheet: Worksheet, anchor: DrawingAnchor): RangeAddress {
	const row: AxisSize = (r) => printRowHeight(sheet, r);
	const col: AxisSize = (c) => printColWidth(sheet, c);
	switch (anchor.type) {
		case 'twoCell': {
			const { from, to } = anchor;
			const right = to.colOff > 0 ? to.col : Math.max(from.col, to.col - 1);
			const bottom = to.rowOff > 0 ? to.row : Math.max(from.row, to.row - 1);
			return { top: from.row, left: from.col, bottom, right };
		}
		case 'oneCell': {
			const { from, ext } = anchor;
			return {
				top: from.row,
				left: from.col,
				bottom: walk(
					row,
					from.row,
					from.rowOff / EMU_PER_POINT,
					ext.cy / EMU_PER_POINT,
					MAX_ROW_INDEX,
				),
				right: walk(
					col,
					from.col,
					from.colOff / EMU_PER_POINT,
					ext.cx / EMU_PER_POINT,
					MAX_COL_INDEX,
				),
			};
		}
		case 'absolute': {
			const x = anchor.pos.x / EMU_PER_POINT;
			const y = anchor.pos.y / EMU_PER_POINT;
			const left = walk(col, 0, 0, x + 0.001, MAX_COL_INDEX);
			const top = walk(row, 0, 0, y + 0.001, MAX_ROW_INDEX);
			return {
				top,
				left,
				bottom: walk(row, 0, 0, y + anchor.ext.cy / EMU_PER_POINT, MAX_ROW_INDEX),
				right: walk(col, 0, 0, x + anchor.ext.cx / EMU_PER_POINT, MAX_COL_INDEX),
			};
		}
	}
}

function visibleFormat(styles: StyleSheet | undefined, xfId: number | undefined): boolean {
	if (!styles || xfId === undefined || xfId <= 0) {
		return false;
	}
	const xf = styles.cellXfs[xfId];
	if (!xf) {
		return false;
	}
	const fill = styles.fills[xf.fillId];
	if (xf.fillId > 1 && fill && (fill.type === 'gradient' || fill.patternType !== 'none')) {
		return true;
	}
	const border = styles.borders[xf.borderId];
	if (border) {
		for (const edge of [border.left, border.right, border.top, border.bottom, border.diagonal]) {
			if (edge && edge.style !== 'none') {
				return true;
			}
		}
	}
	return false;
}

/**
 * The range Excel prints without a print area, or `null` when the sheet has
 * nothing to print. `styles` enables formatting-only cells.
 */
export function printableRange(sheet: Worksheet, styles?: StyleSheet): RangeAddress | null {
	let top = Infinity;
	let left = Infinity;
	let bottom = -1;
	let right = -1;
	for (const [key, cell] of sheet.cells) {
		const used =
			cell.value !== null ||
			cell.formula !== undefined ||
			(cell.richText?.length ?? 0) > 0 ||
			visibleFormat(styles, cell.style);
		if (!used) {
			continue;
		}
		const r = keyRow(key);
		const c = keyCol(key);
		top = Math.min(top, r);
		bottom = Math.max(bottom, r);
		left = Math.min(left, c);
		right = Math.max(right, c);
	}
	let range: RangeAddress | null =
		bottom < 0 ? null : { top, left, bottom: Math.max(top, bottom), right: Math.max(left, right) };
	for (const obj of sheet.drawings) {
		if (obj.hidden) {
			continue;
		}
		const r = anchorCellRange(sheet, obj.anchor);
		range = range ? rangeUnion(range, r) : r;
	}
	if (!range) {
		return null;
	}
	// Merged areas that overlap the edge print whole.
	for (let pass = 0, grown = true; grown && pass < 4; pass++) {
		grown = false;
		for (const m of sheet.merges) {
			if (rangesIntersect(m, range) && !rangeContainsRange(range, m)) {
				range = rangeUnion(range, m);
				grown = true;
			}
		}
	}
	return range;
}

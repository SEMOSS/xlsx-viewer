import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { rangeContainsRange, rangesIntersect } from '../address/range';
import {
	adjustAreaForStructureChange,
	adjustRangeForStructureChange,
	type StructureChange,
} from '../formula/language/adjust-range';
import { isLocalReference, referenceEdit, referenceSheet } from '../formula/language/ref-edits';
import { sameSheetName } from '../formula/language/references';
import { applyEdits, type TextEdit } from '../formula/language/token-utils';
import { tokenizeFormula } from '../formula/language/tokenizer';
import type { RefArea } from '../formula/language/tokens';
import type { CellAddress, RangeAddress } from '../types/address';

/**
 * Insert / delete cells with shifting, described as a structure change
 * limited to a band: the rows (or columns) of the band move along the shift
 * axis, everything outside the band stays where it is.
 */

/** Direction of an Insert Cells (`right`, `down`) or Delete Cells (`left`, `up`) shift. */
export type CellShiftDirection = 'right' | 'down' | 'left' | 'up';

export interface CellBandShift {
	/** Sheet whose cells move. */
	sheet: string;
	/** Cells move up / down (otherwise left / right). */
	vertical: boolean;
	insert: boolean;
	/** First inserted / deleted line along the shift axis (a row for vertical shifts). */
	at: number;
	count: number;
	/** Inclusive extent of the band across the shift axis (columns for vertical shifts). */
	from: number;
	to: number;
}

/** The band shift performed by inserting / deleting `range` with `shift`. */
export function bandShiftOf(
	sheet: string,
	range: RangeAddress,
	shift: CellShiftDirection,
): CellBandShift {
	const vertical = shift === 'down' || shift === 'up';
	const insert = shift === 'down' || shift === 'right';
	return vertical
		? {
				sheet,
				vertical,
				insert,
				at: range.top,
				count: range.bottom - range.top + 1,
				from: range.left,
				to: range.right,
			}
		: {
				sheet,
				vertical,
				insert,
				at: range.left,
				count: range.right - range.left + 1,
				from: range.top,
				to: range.bottom,
			};
}

/** The whole-row / whole-column change the band shift behaves like inside the band. */
export function structureChangeOf(band: CellBandShift): StructureChange {
	let kind: StructureChange['kind'];
	if (band.vertical) {
		kind = band.insert ? 'insertRows' : 'deleteRows';
	} else {
		kind = band.insert ? 'insertCols' : 'deleteCols';
	}
	return { kind, sheet: band.sheet, at: band.at, count: band.count };
}

function maxLine(band: CellBandShift): number {
	return band.vertical ? MAX_ROW_INDEX : MAX_COL_INDEX;
}

/** A cross-axis span lies entirely inside the band. */
function insideBand(lo: number, hi: number, band: CellBandShift): boolean {
	return Math.min(lo, hi) >= band.from && Math.max(lo, hi) <= band.to;
}

/**
 * Adjust a reference area as written. Only references lying entirely inside
 * the band's cross extent change (like a row / column insertion or deletion);
 * whole columns ignore vertical shifts and whole rows horizontal ones.
 * Returns the input when unchanged and `null` when the cells are deleted.
 */
export function adjustAreaForBand(area: RefArea, band: CellBandShift): RefArea | null {
	if (band.vertical ? area.kind === 'columns' : area.kind === 'rows') {
		return area;
	}
	const lo = band.vertical ? area.start.col : area.start.row;
	const hi = band.vertical ? area.end.col : area.end.row;
	if (!insideBand(lo, hi, band)) {
		return area;
	}
	return adjustAreaForStructureChange(area, structureChangeOf(band));
}

/** Same rule for a stored range (merges, formats, validations, tables, links). */
export function adjustRangeForBand(range: RangeAddress, band: CellBandShift): RangeAddress | null {
	const lo = band.vertical ? range.left : range.top;
	const hi = band.vertical ? range.right : range.bottom;
	if (!insideBand(lo, hi, band)) {
		return range;
	}
	return adjustRangeForStructureChange(range, structureChangeOf(band));
}

/** Rewrite the references of a formula living on `formulaSheet`. */
export function adjustFormulaForBand(
	formula: string,
	band: CellBandShift,
	formulaSheet: string,
): string {
	if (formula === '' || band.count <= 0) {
		return formula;
	}
	const tokens = tokenizeFormula(formula);
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if (!isLocalReference(t) || !sameSheetName(referenceSheet(t, formulaSheet), band.sheet)) {
			continue;
		}
		const adjusted = adjustAreaForBand(t.area, band);
		if (adjusted !== t.area) {
			edits.push(referenceEdit(tokens, i, adjusted));
		}
	}
	return edits.length === 0 ? formula : applyEdits(formula, edits);
}

/**
 * Where a cell-attached point (cell, note, comment, sparkline) ends up:
 * the same object when it does not move, `null` when its cell is deleted or
 * pushed off the grid.
 */
export function movePointForBand<T extends CellAddress>(p: T, band: CellBandShift): T | null {
	const line = band.vertical ? p.row : p.col;
	const cross = band.vertical ? p.col : p.row;
	if (cross < band.from || cross > band.to || line < band.at) {
		return p;
	}
	let next: number;
	if (band.insert) {
		next = line + band.count;
		if (next > maxLine(band)) {
			return null;
		}
	} else {
		if (line < band.at + band.count) {
			return null;
		}
		next = line - band.count;
	}
	return band.vertical ? { ...p, row: next } : { ...p, col: next };
}

/** Areas of the sheet touched by a band shift. */
export interface BandZones {
	/** Cells that move (for deletions: the cells after the deleted ones). */
	block: RangeAddress | null;
	/** Deleted cells (deletions only). */
	removed: RangeAddress | null;
	/** Everything affected: the block plus the deleted cells. */
	zone: RangeAddress;
}

function bandRect(band: CellBandShift, lo: number, hi: number): RangeAddress {
	return band.vertical
		? { top: lo, bottom: hi, left: band.from, right: band.to }
		: { top: band.from, bottom: band.to, left: lo, right: hi };
}

export function bandZones(band: CellBandShift): BandZones {
	const max = maxLine(band);
	const zone = bandRect(band, band.at, max);
	if (band.insert) {
		return { block: zone, removed: null, zone };
	}
	const after = band.at + band.count;
	return {
		block: after <= max ? bandRect(band, after, max) : null,
		removed: bandRect(band, band.at, Math.min(max, after - 1)),
		zone,
	};
}

/** How a rectangular object (merge, array, table) fares: kept, moved, deleted or split. */
export type BandFate = 'outside' | 'moves' | 'removed' | 'split';

export function bandFateOf(rect: RangeAddress, zones: BandZones): BandFate {
	if (!rangesIntersect(rect, zones.zone)) {
		return 'outside';
	}
	if (zones.block && rangeContainsRange(zones.block, rect)) {
		return 'moves';
	}
	if (zones.removed && rangeContainsRange(zones.removed, rect)) {
		return 'removed';
	}
	return 'split';
}

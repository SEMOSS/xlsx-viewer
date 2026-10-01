/** Source-range helpers for chart creation: used bounds and Excel's "current region". */
import { cellKey, keyCol, keyRow } from '../../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import type { RangeAddress } from '../../types/address';
import type { Worksheet } from '../../types/sheet';

/** Bounding box of the non-empty cells of a sheet, or `null` when it is empty. */
export function usedBounds(sheet: Worksheet): RangeAddress | null {
	let top = Infinity;
	let left = Infinity;
	let bottom = -1;
	let right = -1;
	for (const [key, cell] of sheet.cells) {
		if (cell.value === null || cell.value === '') {
			continue;
		}
		const r = keyRow(key);
		const c = keyCol(key);
		top = Math.min(top, r);
		left = Math.min(left, c);
		bottom = Math.max(bottom, r);
		right = Math.max(right, c);
	}
	return bottom < 0 ? null : { top, left, bottom, right };
}

function hasValue(sheet: Worksheet, r0: number, c0: number, r1: number, c1: number): boolean {
	for (let r = Math.max(0, r0); r <= Math.min(MAX_ROW_INDEX, r1); r++) {
		for (let c = Math.max(0, c0); c <= Math.min(MAX_COL_INDEX, c1); c++) {
			const v = sheet.cells.get(cellKey(r, c))?.value;
			if (v !== null && v !== undefined && v !== '') {
				return true;
			}
		}
	}
	return false;
}

/** Excel's "current region": the block bounded by empty rows and columns around a cell. */
export function currentRegion(sheet: Worksheet, start: RangeAddress): RangeAddress {
	const r = { ...start };
	for (let guard = 0; guard < 100_000; guard++) {
		if (r.top > 0 && hasValue(sheet, r.top - 1, r.left - 1, r.top - 1, r.right + 1)) {
			r.top--;
		} else if (
			r.bottom < MAX_ROW_INDEX &&
			hasValue(sheet, r.bottom + 1, r.left - 1, r.bottom + 1, r.right + 1)
		) {
			r.bottom++;
		} else if (r.left > 0 && hasValue(sheet, r.top - 1, r.left - 1, r.bottom + 1, r.left - 1)) {
			r.left--;
		} else if (
			r.right < MAX_COL_INDEX &&
			hasValue(sheet, r.top - 1, r.right + 1, r.bottom + 1, r.right + 1)
		) {
			r.right++;
		} else {
			break;
		}
	}
	return r;
}

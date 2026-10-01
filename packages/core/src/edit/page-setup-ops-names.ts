import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { rangeContainsRange } from '../address/range';
import {
	findPrintName,
	formatPrintArea,
	formatPrintTitles,
	getPrintArea,
	PRINT_AREA_NAME,
	PRINT_TITLES_NAME,
	type PrintTitles,
} from '../print/print-names';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { compareDefinedNames } from './names-ops';
import type { EditTransaction } from './transaction';

/**
 * Page Layout > Print Area and Print Titles: the sheet-scoped built-in
 * names `_xlnm.Print_Area` and `_xlnm.Print_Titles`, kept with absolute,
 * sheet-qualified references.
 */

function setPrintName(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	which: 'Print_Area' | 'Print_Titles',
	formula: string | null,
): void {
	tx.touchWorkbook('names');
	const names = workbook.definedNames;
	const current = findPrintName(workbook, sheetIndex, which);
	const at = current ? names.indexOf(current) : -1;
	if (formula === null || formula === '') {
		if (at >= 0) {
			names.splice(at, 1);
		}
	} else if (at >= 0) {
		names[at] = { ...current!, formula };
	} else {
		const def = {
			name: which === 'Print_Area' ? PRINT_AREA_NAME : PRINT_TITLES_NAME,
			formula,
			localSheet: sheetIndex,
			builtin: true,
		};
		const pos = names.findIndex((n) => compareDefinedNames(n, def) > 0);
		names.splice(pos < 0 ? names.length : pos, 0, def);
	}
	workbook.dirty = true;
}

/** Set Print Area (several ranges print on separate pages); `null` or `[]` clears it. */
export function setPrintArea(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[] | null,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const formula = ranges && ranges.length ? formatPrintArea(sheet.name, ranges) : null;
	setPrintName(tx, workbook, sheetIndex, 'Print_Area', formula);
}

export function clearPrintArea(tx: EditTransaction, workbook: Workbook, sheetIndex: number): void {
	setPrintName(tx, workbook, sheetIndex, 'Print_Area', null);
}

/** Two areas that join into one rectangle (same rows and touching columns, or the reverse). */
function joinAreas(a: RangeAddress, b: RangeAddress): RangeAddress | null {
	if (
		a.top === b.top &&
		a.bottom === b.bottom &&
		(a.right + 1 === b.left || b.right + 1 === a.left)
	) {
		return { ...a, left: Math.min(a.left, b.left), right: Math.max(a.right, b.right) };
	}
	if (
		a.left === b.left &&
		a.right === b.right &&
		(a.bottom + 1 === b.top || b.bottom + 1 === a.top)
	) {
		return { ...a, top: Math.min(a.top, b.top), bottom: Math.max(a.bottom, b.bottom) };
	}
	return null;
}

/** Add to Print Area: adjacent ranges extend an area, others become new areas. */
export function addToPrintArea(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
): void {
	const areas = getPrintArea(workbook, sheetIndex);
	for (const r of ranges) {
		if (areas.some((a) => rangeContainsRange(a, r))) {
			continue;
		}
		let next = r;
		for (let i = areas.length - 1; i >= 0; i--) {
			if (rangeContainsRange(next, areas[i])) {
				areas.splice(i, 1);
			}
		}
		const i = areas.findIndex((a) => joinAreas(a, next) !== null);
		if (i >= 0) {
			next = joinAreas(areas[i], next)!;
			areas[i] = next;
		} else {
			areas.push(next);
		}
	}
	setPrintArea(tx, workbook, sheetIndex, areas);
}

/** Rows to repeat at top / columns to repeat at left (`undefined` clears one). */
export function setPrintTitles(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	titles: PrintTitles,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const clamp = (s: PrintTitles['rows'], max: number) =>
		s
			? {
					start: Math.max(0, Math.min(s.start, s.end)),
					end: Math.min(max, Math.max(s.start, s.end)),
				}
			: undefined;
	const formula = formatPrintTitles(sheet.name, {
		rows: clamp(titles.rows, MAX_ROW_INDEX),
		cols: clamp(titles.cols, MAX_COL_INDEX),
	});
	setPrintName(tx, workbook, sheetIndex, 'Print_Titles', formula || null);
}

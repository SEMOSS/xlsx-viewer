import type { CellAddress } from '../types/address';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/**
 * Page Layout > Breaks: manual page breaks above rows (`rowBreaks`) and
 * left of columns (`colBreaks`), and how they follow inserted and deleted
 * rows and columns.
 */

function sheetFor(tx: EditTransaction, workbook: Workbook, sheetIndex: number): Worksheet {
	tx.touchSheet(sheetIndex, false, 'view');
	return workbook.sheets[sheetIndex];
}

/** Which breaks an insert / remove affects: both, or only the row (column) one for whole rows (columns). */
export type PageBreakAxes = 'both' | 'row' | 'col';

function withBreak(list: readonly number[] | undefined, at: number, on: boolean): number[] {
	const set = new Set(list ?? []);
	if (on) {
		set.add(at);
	} else {
		set.delete(at);
	}
	return [...set].sort((a, b) => a - b);
}

/**
 * Insert Page Break at a cell: a break above its row and left of its
 * column (none above row 1 or left of column A).
 */
export function insertPageBreak(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	cell: CellAddress,
	axes: PageBreakAxes = 'both',
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	if (axes !== 'col' && cell.row > 0) {
		sheet.rowBreaks = withBreak(sheet.rowBreaks, cell.row, true);
	}
	if (axes !== 'row' && cell.col > 0) {
		sheet.colBreaks = withBreak(sheet.colBreaks, cell.col, true);
	}
}

/** The cell is next to a manual break that Remove Page Break would delete. */
export function canRemovePageBreak(sheet: Worksheet, cell: CellAddress): boolean {
	return (sheet.rowBreaks ?? []).includes(cell.row) || (sheet.colBreaks ?? []).includes(cell.col);
}

/** Remove Page Break: the manual breaks above the cell's row and left of its column. */
export function removePageBreak(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	cell: CellAddress,
	axes: PageBreakAxes = 'both',
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	if (axes !== 'col' && sheet.rowBreaks?.includes(cell.row)) {
		sheet.rowBreaks = withBreak(sheet.rowBreaks, cell.row, false);
	}
	if (axes !== 'row' && sheet.colBreaks?.includes(cell.col)) {
		sheet.colBreaks = withBreak(sheet.colBreaks, cell.col, false);
	}
}

/** Reset All Page Breaks: every manual break goes. */
export function resetAllPageBreaks(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	sheet.rowBreaks = [];
	sheet.colBreaks = [];
}

/**
 * Manual breaks after inserting (deleting) `count` rows or columns at `at`:
 * breaks move with their row; breaks of deleted rows disappear.
 */
export function shiftPageBreaks(
	breaks: readonly number[] | undefined,
	insert: boolean,
	at: number,
	count: number,
): number[] | undefined {
	if (!breaks) {
		return undefined;
	}
	const out = new Set<number>();
	for (const b of breaks) {
		if (b < at) {
			out.add(b);
		} else if (insert) {
			out.add(b + count);
		} else if (b >= at + count) {
			out.add(b - count);
		}
	}
	return [...out].filter((b) => b > 0).sort((a, b) => a - b);
}

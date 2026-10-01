import { cellKey } from '../address/cell-address';
import { rangesIntersect } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Table } from '../types/table';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

/** Shared helpers of the table edit operations (not part of the public API). */

export function replaceTable(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	next: Table,
): void {
	tx.touchSheet(sheetIndex);
	const sheet = workbook.sheets[sheetIndex];
	sheet.tables = sheet.tables.map((t, i) => (i === index ? { ...next, dirty: true } : t));
	sheet.dirty = true;
}

export function tableOf(workbook: Workbook, sheetIndex: number, index: number): Table {
	const table = workbook.sheets[sheetIndex].tables[index];
	if (!table) {
		throw new RangeError(`No table ${index} on sheet ${sheetIndex}`);
	}
	return table;
}

/** Header row plus data rows: the range the filter buttons cover. */
export function filterRef(table: Pick<Table, 'ref' | 'totalsRowCount'>): RangeAddress {
	return { ...table.ref, bottom: table.ref.bottom - (table.totalsRowCount > 0 ? 1 : 0) };
}

/** The cells of `row` from `left` to `right` are empty and outside other tables and merges. */
export function rowIsFree(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	left: number,
	right: number,
	self: Table,
): boolean {
	const sheet = workbook.sheets[sheetIndex];
	if (row < 0 || row > 1048575) {
		return false;
	}
	const band = { top: row, bottom: row, left, right };
	if (
		sheet.tables.some((t) => t !== self && rangesIntersect(t.ref, band)) ||
		sheet.merges.some((m) => rangesIntersect(m, band))
	) {
		return false;
	}
	for (let c = left; c <= right; c++) {
		const cell = sheet.cells.get(cellKey(row, c));
		if (cell && (cell.value !== null || cell.formula)) {
			return false;
		}
	}
	return true;
}

export function setRowCells(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	left: number,
	values: Array<Pick<Cell, 'value' | 'formula'>>,
): void {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchRange(sheetIndex, { top: row, bottom: row, left, right: left + values.length - 1 });
	values.forEach((v, i) => {
		const key = cellKey(row, left + i);
		const existing = sheet.cells.get(key);
		if (v.value === null && !v.formula && existing?.style === undefined) {
			sheet.cells.delete(key);
		} else {
			sheet.cells.set(key, {
				...(existing ?? {}),
				value: v.value,
				formula: v.formula,
				array: undefined,
				arrayAnchor: undefined,
			});
		}
	});
}

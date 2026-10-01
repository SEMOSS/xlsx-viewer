import { cellKey } from '../address/cell-address';
import { rangesIntersect } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { filterRef, replaceTable, setRowCells } from './table-internal';
import type { EditTransaction } from './transaction';

/**
 * Excel's "Include new rows and columns in table": an entry typed directly
 * below a table (without a totals row) or directly right of it grows the
 * table to include it. Returns true when a table was expanded.
 */
export function autoExpandTable(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): boolean {
	const sheet = workbook.sheets[sheetIndex];
	const index = sheet.tables.findIndex(
		(t) =>
			(row === t.ref.bottom + 1 &&
				t.totalsRowCount === 0 &&
				col >= t.ref.left &&
				col <= t.ref.right) ||
			(col === t.ref.right + 1 && row >= t.ref.top && row <= t.ref.bottom),
	);
	if (index < 0) {
		return false;
	}
	const table = sheet.tables[index];
	const blocked = (band: RangeAddress) =>
		sheet.tables.some((t, i) => i !== index && rangesIntersect(t.ref, band)) ||
		sheet.merges.some((m) => rangesIntersect(m, band));
	if (row === table.ref.bottom + 1) {
		const band = { top: row, bottom: row, left: table.ref.left, right: table.ref.right };
		if (blocked(band)) {
			return false;
		}
		const ref = { ...table.ref, bottom: row };
		replaceTable(tx, workbook, sheetIndex, index, {
			...table,
			ref,
			autoFilter: table.autoFilter
				? { ...table.autoFilter, ref: filterRef({ ref, totalsRowCount: 0 }) }
				: undefined,
		});
		return true;
	}
	const band = { top: table.ref.top, bottom: table.ref.bottom, left: col, right: col };
	if (blocked(band)) {
		return false;
	}
	// New column: its header is the header-row text, or "ColumnN" (written into the header cell).
	const used = new Set(table.columns.map((c) => c.name.toLowerCase()));
	let name = '';
	if (table.headerRowCount > 0) {
		const v = sheet.cells.get(cellKey(table.ref.top, col))?.value;
		name = v === null || v === undefined ? '' : String(v).trim();
	}
	for (let n = table.columns.length + 1; !name || used.has(name.toLowerCase()); n++) {
		name = `Column${n}`;
	}
	if (table.headerRowCount > 0 && sheet.cells.get(cellKey(table.ref.top, col))?.value !== name) {
		setRowCells(tx, workbook, sheetIndex, table.ref.top, col, [{ value: name }]);
	}
	const id = Math.max(0, ...table.columns.map((c) => c.id)) + 1;
	const ref = { ...table.ref, right: col };
	replaceTable(tx, workbook, sheetIndex, index, {
		...table,
		ref,
		columns: [...table.columns, { id, name }],
		autoFilter: table.autoFilter
			? { ...table.autoFilter, ref: filterRef({ ref, totalsRowCount: table.totalsRowCount }) }
			: undefined,
	});
	return true;
}

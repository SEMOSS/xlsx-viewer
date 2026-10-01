import { cellKey } from '../address/cell-address';
import { rangesIntersect } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Table } from '../types/table';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import { insertRows } from './structure-ops';
import type { EditTransaction } from './transaction';

export class TableError extends Error {
	constructor(
		message: string,
		readonly reason: 'overlap' | 'merged' | 'blocked' | 'invalid-name' | 'duplicate',
	) {
		super(message);
		this.name = 'TableError';
	}
}

/** Next unused table id and `TableN` name across the workbook. */
export function nextTableIdentity(workbook: Workbook): { id: number; name: string } {
	let id = 0;
	const names = new Set<string>();
	for (const s of workbook.sheets) {
		for (const t of s.tables) {
			id = Math.max(id, t.id);
			names.add(t.name.toLowerCase());
		}
	}
	for (const n of workbook.definedNames) {
		names.add(n.name.toLowerCase());
	}
	let i = 1;
	while (names.has(`table${i}`)) {
		i++;
	}
	return { id: id + 1, name: `Table${i}` };
}

/**
 * Create an Excel table (Insert > Table / Format as Table) over a range.
 * With `hasHeaders: false` a header row (`Column1`, ...) is inserted above
 * the data, as Excel does. Header names are made unique.
 */
export function createTable(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	options: { hasHeaders?: boolean; styleName?: string } = {},
): Table {
	const sheet = workbook.sheets[sheetIndex];
	if (sheet.tables.some((t) => rangesIntersect(t.ref, range))) {
		throw new TableError("A table can't overlap another table.", 'overlap');
	}
	if (sheet.merges.some((m) => rangesIntersect(m, range))) {
		throw new TableError("A table can't overlap a range that contains merged cells.", 'merged');
	}
	let ref = range;
	if (options.hasHeaders === false) {
		insertRows(tx, workbook, sheetIndex, range.top, 1);
		ref = { ...range, bottom: range.bottom + 1 };
	}
	tx.touchSheet(sheetIndex);
	const used = new Set<string>();
	const columns = [];
	for (let c = ref.left; c <= ref.right; c++) {
		const key = cellKey(ref.top, c);
		const v = sheet.cells.get(key)?.value;
		let name = v === null || v === undefined || isCellError(v) ? '' : String(v).trim();
		if (!name || options.hasHeaders === false) {
			name = `Column${c - ref.left + 1}`;
		}
		let unique = name;
		let n = 2;
		while (used.has(unique.toLowerCase())) {
			unique = `${name}${n++}`;
		}
		used.add(unique.toLowerCase());
		const existing = sheet.cells.get(key);
		if (existing?.value !== unique) {
			sheet.cells.set(key, { ...(existing ?? {}), value: unique, formula: undefined });
			tx.markChanged({ sheet: sheetIndex, row: ref.top, col: c });
		}
		columns.push({ id: c - ref.left + 1, name: unique });
	}
	const identity = nextTableIdentity(workbook);
	const table: Table = {
		id: identity.id,
		name: identity.name,
		displayName: identity.name,
		ref,
		headerRowCount: 1,
		totalsRowCount: 0,
		columns,
		autoFilter: { ref, columns: [] },
		style: {
			name: options.styleName ?? workbook.styles.defaultTableStyle ?? 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: false,
			showRowStripes: true,
			showColumnStripes: false,
		},
		dirty: true,
	};
	sheet.tables = [...sheet.tables, table];
	if (sheet.autoFilter && rangesIntersect(sheet.autoFilter.ref, ref)) {
		sheet.autoFilter = undefined;
	}
	sheet.dirty = true;
	return table;
}

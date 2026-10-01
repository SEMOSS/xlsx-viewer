import type { RangeAddress } from '../../types/address';
import type { CellError } from '../../types/cell';
import type { Table, TableColumn } from '../../types/table';
import { ERRORS } from '../../values/errors';
import type { StructuredRefNode } from '../ast';

/**
 * Resolution of structured (table) references to sheet ranges. Shared by
 * the evaluator and by the dependency analysis of the calculation engine.
 */

const columnMaps = new WeakMap<readonly TableColumn[], Map<string, number>>();

/** Index of a table column by name (case-insensitive), -1 when missing. */
export function tableColumnIndex(table: Table, name: string): number {
	let map = columnMaps.get(table.columns);
	if (map === undefined) {
		map = new Map();
		table.columns.forEach((c, i) => {
			const key = c.name.toLowerCase();
			if (!map!.has(key)) {
				map!.set(key, i);
			}
		});
		columnMaps.set(table.columns, map);
	}
	return map.get(name.toLowerCase()) ?? -1;
}

/** Data body rows of a table (may be empty: `top > bottom`). */
export function tableDataRows(table: Table): { top: number; bottom: number } {
	return {
		top: table.ref.top + (table.headerRowCount > 0 ? 1 : 0),
		bottom: table.ref.bottom - (table.totalsRowCount > 0 ? 1 : 0),
	};
}

type Parts = Pick<StructuredRefNode, 'specifiers' | 'columnStart' | 'columnEnd'>;

function rowsOf(
	parts: Parts,
	table: Table,
	row: number,
): { top: number; bottom: number } | CellError {
	const ref = table.ref;
	const data = tableDataRows(table);
	const specs = parts.specifiers;
	if (specs.length === 0) {
		return data.top <= data.bottom ? data : ERRORS.REF;
	}
	if (specs.includes('#This Row')) {
		if (specs.length > 1 || row < data.top || row > data.bottom) {
			return ERRORS.VALUE;
		}
		return { top: row, bottom: row };
	}
	if (specs.includes('#All')) {
		return { top: ref.top, bottom: ref.bottom };
	}
	const headers = specs.includes('#Headers');
	const body = specs.includes('#Data');
	const totals = specs.includes('#Totals');
	if ((headers && table.headerRowCount <= 0) || (totals && table.totalsRowCount <= 0)) {
		return ERRORS.REF;
	}
	if (headers && totals && !body) {
		return ERRORS.REF;
	}
	const top = headers ? ref.top : body ? data.top : ref.bottom;
	const bottom = totals ? ref.bottom : body ? data.bottom : ref.top;
	return top <= bottom ? { top, bottom } : ERRORS.REF;
}

/**
 * Range addressed by a structured reference within `table`, seen from a
 * formula on `row` (for `[@Col]` / `[#This Row]`). `#REF!` for unknown
 * columns or missing table parts, `#VALUE!` for a "this row" reference
 * outside the table's data rows.
 */
export function structuredRange(parts: Parts, table: Table, row: number): RangeAddress | CellError {
	const rows = rowsOf(parts, table, row);
	if ('code' in rows) {
		return rows;
	}
	let left = table.ref.left;
	let right = table.ref.right;
	if (parts.columnStart !== undefined) {
		const i = tableColumnIndex(table, parts.columnStart);
		const j = parts.columnEnd === undefined ? i : tableColumnIndex(table, parts.columnEnd);
		if (i < 0 || j < 0) {
			return ERRORS.REF;
		}
		left = table.ref.left + Math.min(i, j);
		right = table.ref.left + Math.max(i, j);
		if (right > table.ref.right) {
			return ERRORS.REF;
		}
	}
	return { top: rows.top, left, bottom: rows.bottom, right };
}

/** Range of a bare table name (`=SUM(Table1)`): the data body, all columns. */
export function tableDataRange(table: Table): RangeAddress | CellError {
	return structuredRange({ specifiers: [] }, table, -1);
}

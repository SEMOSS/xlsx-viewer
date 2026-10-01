import { cellKey } from '../address/cell-address';
import { displayText } from '../convert/text-grid';
import { wildcardToRegExp } from '../formula/criteria';
import { StyleResolver } from '../styles/resolve-style';
import type { RangeAddress } from '../types/address';
import type { AutoFilter, CustomFilter, FilterColumn } from '../types/filter';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import type { EditTransaction } from './transaction';

/** Where a filter lives: the sheet's AutoFilter or a table's. */
export type FilterTarget = { kind: 'sheet' } | { kind: 'table'; index: number };

export function getFilter(sheet: Worksheet, target: FilterTarget): AutoFilter | undefined {
	if (target.kind === 'sheet') {
		return sheet.autoFilter;
	}
	const table = sheet.tables[target.index];
	if (!table) {
		return undefined;
	}
	return (
		table.autoFilter ??
		(table.headerRowCount
			? {
					ref: { ...table.ref, bottom: table.ref.bottom - (table.totalsRowCount ? 1 : 0) },
					columns: [],
				}
			: undefined)
	);
}

function setFilter(sheet: Worksheet, target: FilterTarget, filter: AutoFilter): void {
	if (target.kind === 'sheet') {
		sheet.autoFilter = filter;
	} else {
		const table = sheet.tables[target.index];
		sheet.tables[target.index] = { ...table, autoFilter: filter, dirty: true };
	}
}

function matchCustom(text: string, value: unknown, f: CustomFilter): boolean {
	const op = f.operator ?? 'equal';
	const num = Number(f.val);
	if (typeof value === 'number' && f.val.trim() !== '' && Number.isFinite(num)) {
		switch (op) {
			case 'equal':
				return value === num;
			case 'notEqual':
				return value !== num;
			case 'lessThan':
				return value < num;
			case 'lessThanOrEqual':
				return value <= num;
			case 'greaterThan':
				return value > num;
			case 'greaterThanOrEqual':
				return value >= num;
		}
	}
	if (op === 'equal' || op === 'notEqual') {
		const hit = wildcardToRegExp(f.val).test(text);
		return op === 'equal' ? hit : !hit;
	}
	const c = text.localeCompare(f.val, undefined, { sensitivity: 'base' });
	switch (op) {
		case 'lessThan':
			return c < 0;
		case 'lessThanOrEqual':
			return c <= 0;
		case 'greaterThan':
			return c > 0;
		default:
			return c >= 0;
	}
}

/** Whether a cell passes one filter column's criteria. */
export function passesFilterColumn(
	column: FilterColumn,
	text: string,
	value: unknown,
	stats?: { sortedDesc: number[] },
): boolean {
	if (column.filters) {
		if (text === '') {
			return column.filters.blank === true;
		}
		const lower = text.toLowerCase();
		return column.filters.values.some((v) => v.toLowerCase() === lower);
	}
	if (column.customFilters) {
		const results = column.customFilters.filters.map((f) => matchCustom(text, value, f));
		return column.customFilters.and ? results.every(Boolean) : results.some(Boolean);
	}
	if (column.top10 && typeof value === 'number' && stats) {
		const n = column.top10.percent
			? Math.max(1, Math.floor((stats.sortedDesc.length * column.top10.val) / 100))
			: column.top10.val;
		const list = column.top10.top === false ? [...stats.sortedDesc].reverse() : stats.sortedDesc;
		const threshold = list[Math.min(list.length, n) - 1];
		return column.top10.top === false ? value <= threshold : value >= threshold;
	}
	if (column.top10) {
		return false;
	}
	return true;
}

/**
 * Re-apply every criterion of a filter: rows of the data area (below the
 * header row) that fail any column are hidden, the others shown.
 */
export function applyFilter(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	target: FilterTarget,
): { shown: number; hidden: number } {
	const sheet = workbook.sheets[sheetIndex];
	const filter = getFilter(sheet, target);
	if (!filter) {
		return { shown: 0, hidden: 0 };
	}
	tx.touchSheet(sheetIndex, false, 'filter');
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const ref = filter.ref;
	const stats = new Map<number, { sortedDesc: number[] }>();
	for (const col of filter.columns) {
		if (col.top10) {
			const nums: number[] = [];
			for (let r = ref.top + 1; r <= ref.bottom; r++) {
				const v = sheet.cells.get(cellKey(r, ref.left + col.colId))?.value;
				if (typeof v === 'number') {
					nums.push(v);
				}
			}
			stats.set(col.colId, { sortedDesc: nums.sort((a, b) => b - a) });
		}
	}
	let shown = 0;
	let hidden = 0;
	for (let r = ref.top + 1; r <= ref.bottom; r++) {
		let pass = true;
		for (const col of filter.columns) {
			const cell = sheet.cells.get(cellKey(r, ref.left + col.colId));
			const text = displayText(cell, styles, workbook.date1904);
			const value = cell?.value;
			if (
				!passesFilterColumn(
					col,
					text,
					isCellError(value) ? value.code : value,
					stats.get(col.colId),
				)
			) {
				pass = false;
				break;
			}
		}
		const info = sheet.rows.get(r) ?? {};
		if (pass) {
			shown++;
			if (info.hidden) {
				const { hidden: _h, ...rest } = info;
				if (Object.keys(rest).length) {
					sheet.rows.set(r, rest);
				} else {
					sheet.rows.delete(r);
				}
			}
		} else {
			hidden++;
			sheet.rows.set(r, { ...info, hidden: true });
		}
	}
	sheet.dirty = true;
	return { shown, hidden };
}

/** Set (or clear with `null`) the criteria of one column and re-apply the filter. */
export function setFilterColumn(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	target: FilterTarget,
	colId: number,
	column: Omit<FilterColumn, 'colId'> | null,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const filter = getFilter(sheet, target);
	if (!filter) {
		return;
	}
	tx.touchSheet(sheetIndex, false, 'filter');
	const columns = filter.columns.filter((c) => c.colId !== colId);
	if (column) {
		columns.push({ ...column, colId });
		columns.sort((a, b) => a.colId - b.colId);
	}
	setFilter(sheet, target, { ...filter, columns });
	applyFilter(tx, workbook, sheetIndex, target);
}

/** Distinct displayed values (and whether blanks exist) in a filter column's data area. */
export function filterColumnValues(
	workbook: Workbook,
	sheet: Worksheet,
	ref: RangeAddress,
	colId: number,
): { values: string[]; hasBlank: boolean } {
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const seen = new Map<string, { text: string; sort: number | string }>();
	let hasBlank = false;
	for (let r = ref.top + 1; r <= ref.bottom; r++) {
		const cell = sheet.cells.get(cellKey(r, ref.left + colId));
		const text = displayText(cell, styles, workbook.date1904);
		if (text === '') {
			hasBlank = true;
			continue;
		}
		const key = text.toLowerCase();
		if (!seen.has(key)) {
			seen.set(key, {
				text,
				sort: typeof cell?.value === 'number' ? cell.value : text.toLowerCase(),
			});
		}
	}
	const values = [...seen.values()]
		.sort((a, b) =>
			typeof a.sort === 'number' && typeof b.sort === 'number'
				? a.sort - b.sort
				: typeof a.sort === 'number'
					? -1
					: typeof b.sort === 'number'
						? 1
						: String(a.sort).localeCompare(String(b.sort)),
		)
		.map((v) => v.text);
	return { values, hasBlank };
}

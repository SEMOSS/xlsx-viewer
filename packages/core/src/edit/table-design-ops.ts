import { cellKey } from '../address/cell-address';
import { tableDataRows } from '../formula/evaluator/structured';
import type { Table, TableColumn, TableStyleInfo, TotalsRowFunction } from '../types/table';
import type { Workbook } from '../types/workbook';
import { applyFilter } from './filter-ops';
import { filterRef, replaceTable, rowIsFree, setRowCells, tableOf } from './table-internal';
import { TableError } from './table-ops';
import type { EditTransaction } from './transaction';

const DEFAULT_STYLE_INFO: TableStyleInfo = {
	showFirstColumn: false,
	showLastColumn: false,
	showRowStripes: true,
	showColumnStripes: false,
};

/** SUBTOTAL function numbers Excel uses for totals-row functions (ignoring hidden rows). */
const SUBTOTAL_CODES: Partial<Record<TotalsRowFunction, number>> = {
	average: 101,
	countNums: 102,
	count: 103,
	max: 104,
	min: 105,
	stdDev: 107,
	sum: 109,
	var: 110,
};

/** Index of the table containing a cell of a sheet, or -1. */
export function tableIndexAt(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): number {
	return workbook.sheets[sheetIndex].tables.findIndex(
		(t) => row >= t.ref.top && row <= t.ref.bottom && col >= t.ref.left && col <= t.ref.right,
	);
}

/** Escape a column name for use inside `[...]` of a structured reference. */
export function escapeTableColumnName(name: string): string {
	return name.replace(/(['[\]#])/g, "'$1");
}

/** Totals-row formula for a column function (`SUBTOTAL(109,Table1[Sales])`). */
export function totalsRowFormula(table: Table, column: TableColumn): string | undefined {
	if (column.totalsRowFunction === 'custom') {
		return column.totalsRowFormula;
	}
	const code = column.totalsRowFunction ? SUBTOTAL_CODES[column.totalsRowFunction] : undefined;
	return code
		? `SUBTOTAL(${code},${table.name}[${escapeTableColumnName(column.name)}])`
		: undefined;
}

/** Table Design > Table Styles: switch the table style (built-in or custom name). */
export function setTableStyle(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	styleName: string,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	// An empty name is the gallery's "None": no table style at all.
	replaceTable(tx, workbook, sheetIndex, index, {
		...table,
		style: { ...DEFAULT_STYLE_INFO, ...table.style, name: styleName || undefined },
	});
}

/** Table Design > Table Style Options: banded rows / columns, first / last column. */
export function setTableStyleOptions(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	patch: Partial<Omit<TableStyleInfo, 'name'>>,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	replaceTable(tx, workbook, sheetIndex, index, {
		...table,
		style: { ...DEFAULT_STYLE_INFO, ...table.style, ...patch },
	});
}

/** Remove the criteria of a table's filter and show the rows it hid. */
function clearTableFilter(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	if (table.autoFilter && table.autoFilter.columns.length > 0) {
		replaceTable(tx, workbook, sheetIndex, index, {
			...table,
			autoFilter: { ...table.autoFilter, columns: [] },
		});
		applyFilter(tx, workbook, sheetIndex, { kind: 'table', index });
	}
}

/** Table Design > Filter Button. Needs a header row; hiding clears the table's filters. */
export function setTableFilterButton(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	show: boolean,
): void {
	if (!show) {
		clearTableFilter(tx, workbook, sheetIndex, index);
	}
	const table = tableOf(workbook, sheetIndex, index);
	if (show && table.headerRowCount <= 0) {
		return;
	}
	replaceTable(tx, workbook, sheetIndex, index, {
		...table,
		autoFilter: show ? { ref: filterRef(table), columns: [] } : undefined,
	});
}

/**
 * Table Design > Header Row. Hiding clears the header cells and shrinks
 * the table; showing writes the column names into the row above, which
 * must be empty.
 */
export function setTableHeaderRow(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	show: boolean,
): void {
	let table = tableOf(workbook, sheetIndex, index);
	if (show === table.headerRowCount > 0) {
		return;
	}
	const blanks = table.columns.map(() => ({ value: null }));
	if (!show) {
		clearTableFilter(tx, workbook, sheetIndex, index);
		table = tableOf(workbook, sheetIndex, index);
		setRowCells(tx, workbook, sheetIndex, table.ref.top, table.ref.left, blanks);
		replaceTable(tx, workbook, sheetIndex, index, {
			...table,
			ref: { ...table.ref, top: table.ref.top + 1 },
			headerRowCount: 0,
			autoFilter: undefined,
		});
		return;
	}
	const row = table.ref.top - 1;
	if (!rowIsFree(workbook, sheetIndex, row, table.ref.left, table.ref.right, table)) {
		throw new TableError('There is no empty row above the table for its header row.', 'blocked');
	}
	setRowCells(
		tx,
		workbook,
		sheetIndex,
		row,
		table.ref.left,
		table.columns.map((c) => ({ value: c.name })),
	);
	const next = { ...table, ref: { ...table.ref, top: row }, headerRowCount: 1 };
	replaceTable(tx, workbook, sheetIndex, index, {
		...next,
		autoFilter: { ref: filterRef(next), columns: [] },
	});
}

/**
 * A column holds numbers in its data rows (decides the default totals
 * function). Formula cells not calculated yet count as numbers.
 */
function isNumericColumn(
	workbook: Workbook,
	sheetIndex: number,
	table: Table,
	col: number,
): boolean {
	const sheet = workbook.sheets[sheetIndex];
	const { top, bottom } = tableDataRows(table);
	let formulas = false;
	for (let r = top; r <= bottom; r++) {
		const cell = sheet.cells.get(cellKey(r, col));
		if (cell?.value !== null && cell?.value !== undefined) {
			return typeof cell.value === 'number';
		}
		formulas ||= cell?.formula !== undefined;
	}
	return formulas;
}

/**
 * Table Design > Total Row. Showing adds a totals row below the data (the
 * row must be empty): "Total" in the first column and a SUBTOTAL in the
 * last (SUM for numbers, COUNT otherwise), unless functions were chosen before.
 */
export function setTableTotalsRow(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	show: boolean,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	if (show === table.totalsRowCount > 0) {
		return;
	}
	if (!show) {
		setRowCells(
			tx,
			workbook,
			sheetIndex,
			table.ref.bottom,
			table.ref.left,
			table.columns.map(() => ({ value: null })),
		);
		replaceTable(tx, workbook, sheetIndex, index, {
			...table,
			ref: { ...table.ref, bottom: table.ref.bottom - 1 },
			totalsRowCount: 0,
			totalsRowShown: false,
		});
		return;
	}
	const row = table.ref.bottom + 1;
	if (!rowIsFree(workbook, sheetIndex, row, table.ref.left, table.ref.right, table)) {
		throw new TableError('There is no empty row below the table for its total row.', 'blocked');
	}
	const hasChoices = table.columns.some((c) => c.totalsRowFunction || c.totalsRowLabel);
	const last = table.columns.length - 1;
	const columns = table.columns.map((c, i): TableColumn => {
		if (hasChoices) {
			return c;
		}
		if (i === 0 && last > 0) {
			return { ...c, totalsRowLabel: 'Total' };
		}
		if (i === last) {
			return {
				...c,
				totalsRowFunction: isNumericColumn(workbook, sheetIndex, table, table.ref.left + i)
					? 'sum'
					: 'count',
			};
		}
		return c;
	});
	const next: Table = {
		...table,
		columns,
		ref: { ...table.ref, bottom: row },
		totalsRowCount: 1,
		totalsRowShown: true,
	};
	setRowCells(
		tx,
		workbook,
		sheetIndex,
		row,
		table.ref.left,
		columns.map((c) => {
			const formula = totalsRowFormula(next, c);
			return formula ? { value: null, formula } : { value: c.totalsRowLabel ?? null };
		}),
	);
	replaceTable(tx, workbook, sheetIndex, index, next);
}

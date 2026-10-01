import type { EditTransaction, Table, TableStyleInfo, Workbook } from 'xlsx-viewer-core';
import {
	convertTableToRange,
	createTable,
	renameTable,
	setTableFilterButton,
	setTableHeaderRow,
	setTableStyle,
	setTableStyleOptions,
	setTableTotalsRow,
	TableError,
	tableIndexAt,
} from 'xlsx-viewer-core';

import type { TranslationKey } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';
import { dataBlock } from './structure-commands';

/** The table containing the active cell, if any. */
export function activeTable(store: ViewerStore): { index: number; table: Table } | null {
	const { row, col } = store.selectionOf().active;
	const index = tableIndexAt(store.workbook, store.state.activeSheet, row, col);
	return index < 0 ? null : { index, table: store.activeSheet.tables[index] };
}

/** Run a table edit, reporting TableError messages in the status bar. Returns false on failure. */
function tableEdit(
	store: ViewerStore,
	label: string,
	fn: (tx: EditTransaction, wb: Workbook, sheetIndex: number) => void,
): boolean {
	const sheetIndex = store.state.activeSheet;
	try {
		return store.transact(label, (tx) => fn(tx, store.workbook, sheetIndex));
	} catch (e) {
		if (e instanceof TableError) {
			store.setUi({ message: e.message });
			return false;
		}
		throw e;
	}
}

/** Insert > Table (Ctrl+T): make a table of the selection or the current region. */
export function insertTableCmd(store: ViewerStore, styleName?: string): boolean {
	const block = dataBlock(store);
	// Without the Create Table dialog the first row is taken as the header row (Excel's usual
	// guess); inserting a "Column1..." row instead would move the user's data.
	const hasHeaders = true;
	return tableEdit(store, store.t('undo.createTable'), (tx, wb, si) => {
		createTable(tx, wb, si, block, { hasHeaders, styleName });
	});
}

/**
 * Home > Format as Table: restyle the table at the active cell, or turn the
 * selection (or current region) into a table with the chosen style.
 */
export function formatAsTable(store: ViewerStore, styleName: string): boolean {
	const current = activeTable(store);
	if (current) {
		return tableEdit(store, store.t('undo.tableStyle'), (tx, wb, si) =>
			setTableStyle(tx, wb, si, current.index, styleName),
		);
	}
	return insertTableCmd(store, styleName);
}

/** Table Design > Table Styles. An empty name means "None". */
export function setActiveTableStyle(store: ViewerStore, styleName: string): boolean {
	const current = activeTable(store);
	return (
		!!current &&
		tableEdit(store, store.t('undo.tableStyle'), (tx, wb, si) =>
			setTableStyle(tx, wb, si, current.index, styleName),
		)
	);
}

export type TableOption =
	| 'headerRow'
	| 'totalRow'
	| 'filterButton'
	| keyof Omit<TableStyleInfo, 'name'>;

/** Current state of a Table Style Options checkbox. */
export function tableOptionValue(table: Table, option: TableOption): boolean {
	switch (option) {
		case 'headerRow':
			return table.headerRowCount > 0;
		case 'totalRow':
			return table.totalsRowCount > 0;
		case 'filterButton':
			return table.autoFilter !== undefined;
		default:
			return table.style?.[option] ?? option === 'showRowStripes';
	}
}

/** Undo labels of the Table Style Options checkboxes (the checkbox captions). */
const OPTION_LABELS: Record<TableOption, TranslationKey> = {
	headerRow: 'table.headerRow',
	totalRow: 'table.totalRow',
	filterButton: 'table.filterButton',
	showRowStripes: 'table.bandedRows',
	showColumnStripes: 'table.bandedColumns',
	showFirstColumn: 'table.firstColumn',
	showLastColumn: 'table.lastColumn',
};

/** Table Design > Table Style Options: toggle one checkbox. */
export function setTableOption(store: ViewerStore, option: TableOption, value: boolean): boolean {
	const current = activeTable(store);
	if (!current) {
		return false;
	}
	return tableEdit(store, store.t(OPTION_LABELS[option]), (tx, wb, si) => {
		switch (option) {
			case 'headerRow':
				setTableHeaderRow(tx, wb, si, current.index, value);
				break;
			case 'totalRow':
				setTableTotalsRow(tx, wb, si, current.index, value);
				break;
			case 'filterButton':
				setTableFilterButton(tx, wb, si, current.index, value);
				break;
			default:
				setTableStyleOptions(tx, wb, si, current.index, { [option]: value });
		}
	});
}

/** Table Design > Table Name. Returns an error message, or null when renamed. */
export function renameActiveTable(store: ViewerStore, name: string): string | null {
	const current = activeTable(store);
	if (!current) {
		return null;
	}
	try {
		store.transact(store.t('undo.renameTable'), (tx) =>
			renameTable(tx, store.workbook, store.state.activeSheet, current.index, name),
		);
		return null;
	} catch (e) {
		if (e instanceof TableError) {
			return e.message;
		}
		throw e;
	}
}

/** Table Design > Convert to Range. */
export function convertActiveTableToRange(store: ViewerStore): boolean {
	const current = activeTable(store);
	return (
		!!current &&
		tableEdit(store, store.t('table.convertToRange'), (tx, wb, si) =>
			convertTableToRange(tx, wb, si, current.index),
		)
	);
}

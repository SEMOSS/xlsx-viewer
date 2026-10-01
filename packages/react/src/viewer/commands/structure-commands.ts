import {
	cellKey,
	deleteColumns,
	deleteRows,
	guessHeader,
	insertColumns,
	insertRows,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	SortError,
	sortRange,
} from 'xlsx-viewer-core';

import { currentRegion } from '../../engine/selection/navigation';
import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';
import { probeFor } from './selection-commands';

/** Insert whole rows / columns at the selection (as many as are selected). */
export function insertCellsCmd(store: ViewerStore, axis: 'rows' | 'cols'): void {
	const sheetIndex = store.state.activeSheet;
	const ranges = [...store.selectionOf().ranges].sort((a, b) =>
		axis === 'rows' ? b.top - a.top : b.left - a.left,
	);
	store.transact(store.t(axis === 'rows' ? 'undo.insertRows' : 'undo.insertColumns'), (tx) => {
		for (const r of ranges) {
			if (axis === 'rows') {
				insertRows(tx, store.workbook, sheetIndex, r.top, Math.min(r.bottom - r.top + 1, 10000));
			} else {
				insertColumns(tx, store.workbook, sheetIndex, r.left, Math.min(r.right - r.left + 1, 1000));
			}
		}
	});
}

/** Delete the rows / columns of the selection. */
export function deleteCellsCmd(store: ViewerStore, axis: 'rows' | 'cols'): void {
	const sheetIndex = store.state.activeSheet;
	const ranges = [...store.selectionOf().ranges].sort((a, b) =>
		axis === 'rows' ? b.top - a.top : b.left - a.left,
	);
	store.transact(store.t(axis === 'rows' ? 'undo.deleteRows' : 'undo.deleteColumns'), (tx) => {
		for (const r of ranges) {
			if (axis === 'rows') {
				deleteRows(
					tx,
					store.workbook,
					sheetIndex,
					r.top,
					Math.min(r.bottom - r.top + 1, MAX_ROW_INDEX + 1),
				);
			} else {
				deleteColumns(
					tx,
					store.workbook,
					sheetIndex,
					r.left,
					Math.min(r.right - r.left + 1, MAX_COL_INDEX + 1),
				);
			}
		}
	});
}

/** Ctrl+- / Ctrl+Shift+= with whole rows or columns selected: no dialog needed. */
export function wholeAxisOfSelection(store: ViewerStore): 'rows' | 'cols' | null {
	const r = currentRange(store.selectionOf());
	if (r.left === 0 && r.right === MAX_COL_INDEX) {
		return 'rows';
	}
	if (r.top === 0 && r.bottom === MAX_ROW_INDEX) {
		return 'cols';
	}
	return null;
}

/** The block a sort or filter applies to: the selection, or the current region of a single cell. */
export function dataBlock(store: ViewerStore) {
	const sel = store.selectionOf();
	const r = currentRange(sel);
	if (r.top === r.bottom && r.left === r.right) {
		return currentRegion(probeFor(store.activeSheet), sel.active);
	}
	return r;
}

/** Sort A to Z / Z to A by the active cell's column. */
export function sortSelection(store: ViewerStore, descending: boolean): void {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const block = dataBlock(store);
	const col = store.selectionOf().active.col;
	const hasHeader = guessHeader(sheet, block);
	try {
		store.transact(store.t(descending ? 'cmd.sortZA' : 'cmd.sortAZ'), (tx) => {
			sortRange(tx, store.workbook, sheetIndex, block, [{ index: col, descending }], { hasHeader });
		});
	} catch (e) {
		if (e instanceof SortError) {
			store.setUi({ message: e.message });
			return;
		}
		throw e;
	}
}

/** Data > Filter: add or remove the AutoFilter on the current region. */
export function toggleAutoFilter(store: ViewerStore): void {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	store.transact(store.t('cmd.filter'), (tx) => {
		tx.touchSheet(sheetIndex);
		if (sheet.autoFilter) {
			// Removing the filter shows every hidden row again.
			for (const [row, info] of sheet.rows) {
				if (info.hidden && row >= sheet.autoFilter.ref.top && row <= sheet.autoFilter.ref.bottom) {
					sheet.rows.set(row, { ...info, hidden: undefined });
				}
			}
			sheet.autoFilter = undefined;
		} else {
			const block = dataBlock(store);
			const hasData = sheet.cells.get(cellKey(block.top, block.left)) !== undefined;
			if (!hasData && block.top === block.bottom) {
				return;
			}
			sheet.autoFilter = { ref: block, columns: [] };
		}
		sheet.dirty = true;
	});
}

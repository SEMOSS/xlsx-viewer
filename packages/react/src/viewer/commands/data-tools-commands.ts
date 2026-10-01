import type {
	RangeAddress,
	RemoveDuplicatesOptions,
	RemoveDuplicatesResult,
	TextToColumnsOptions,
} from 'xlsx-viewer-core';
import {
	findDuplicateRows,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	parseTextToColumns,
	RemoveDuplicatesError,
	removeDuplicateRows,
	TextToColumnsError,
	textToColumns,
	textToColumnsLines,
	trimRangeToUsed,
} from 'xlsx-viewer-core';

import type { SelectionState } from '../../engine/selection/selection';
import type { Translate } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';

/** Data > Data Tools: Remove Duplicates and Text to Columns. */

/** Excel's summary after removing duplicates, in the language of `t`. */
export function removeDuplicatesMessage(result: RemoveDuplicatesResult, t: Translate): string {
	if (result.removed === 0) {
		return t('dialog.removeDuplicates.noDuplicates');
	}
	const removed =
		result.removed === 1
			? t('dialog.removeDuplicates.removedOne')
			: t('dialog.removeDuplicates.removedOther', { count: result.removed });
	const remaining =
		result.remaining === 1
			? t('dialog.removeDuplicates.remainOne')
			: t('dialog.removeDuplicates.remainOther', { count: result.remaining });
	return t('dialog.removeDuplicates.summary', { removed, remaining });
}

function selectionOf(range: RangeAddress): SelectionState {
	const active = { row: range.top, col: range.left };
	return { active, anchor: active, ranges: [range] };
}

/**
 * Remove the duplicate rows of `range` as one undoable step, select the
 * remaining block and show Excel's summary in the status bar. Returns the
 * message shown (also for refusals), or null in read-only mode.
 */
export function removeDuplicatesCmd(
	store: ViewerStore,
	range: RangeAddress,
	options: RemoveDuplicatesOptions,
): string | null {
	if (store.readOnly) {
		return null;
	}
	const sheetIndex = store.state.activeSheet;
	let found: ReturnType<typeof findDuplicateRows>;
	try {
		found = findDuplicateRows(store.workbook, sheetIndex, range, options);
	} catch (e) {
		if (e instanceof RemoveDuplicatesError) {
			store.setUi({ message: e.message });
			return e.message;
		}
		throw e;
	}
	let result: RemoveDuplicatesResult = { removed: 0, remaining: found.kept.length };
	if (found.duplicates.length > 0) {
		const area = found.range;
		const kept = {
			...area,
			bottom: area.top + (options.hasHeader ? 1 : 0) + found.kept.length - 1,
		};
		store.transact(
			store.t('cmd.removeDuplicates'),
			(tx) => {
				result = removeDuplicateRows(tx, store.workbook, sheetIndex, range, options);
			},
			{ selectionAfter: selectionOf(kept) },
		);
	}
	const message = removeDuplicatesMessage(result, store.t);
	store.setUi({ message });
	return message;
}

/** Widest converted row, not counting skipped columns (0 when nothing would be written). */
function outputColumns(
	store: ViewerStore,
	source: RangeAddress,
	options: TextToColumnsOptions,
): { columns: number; rows: number } {
	const lines = textToColumnsLines(store.workbook, store.state.activeSheet, source);
	const formats = options.columnFormats ?? [];
	let columns = 0;
	for (const pieces of parseTextToColumns(lines, options)) {
		columns = Math.max(columns, pieces.filter((_, j) => formats[j] !== 'skip').length);
	}
	return { columns, rows: lines.length };
}

/** The cells a conversion fills: the source rows by the widest converted row. */
export function textToColumnsResultRange(
	store: ViewerStore,
	source: RangeAddress,
	options: TextToColumnsOptions,
): RangeAddress {
	const { columns, rows } = outputColumns(store, source, options);
	const { row, col } = options.destination;
	return {
		top: row,
		left: col,
		bottom: Math.min(MAX_ROW_INDEX, row + Math.max(1, rows) - 1),
		right: Math.min(MAX_COL_INDEX, col + Math.max(1, columns) - 1),
	};
}

/**
 * Convert the one-column `source` with Text to Columns as one undoable step
 * and select the result. Returns Excel's message when the conversion is
 * refused (nothing changes), otherwise null.
 */
export function textToColumnsCmd(
	store: ViewerStore,
	source: RangeAddress,
	options: TextToColumnsOptions,
): string | null {
	if (store.readOnly) {
		return null;
	}
	const sheetIndex = store.state.activeSheet;
	const area = trimRangeToUsed(store.activeSheet, source);
	if (outputColumns(store, area, options).columns === 0) {
		// Every column is skipped: nothing to convert.
		return null;
	}
	const after = selectionOf(textToColumnsResultRange(store, area, options));
	try {
		store.transact(
			store.t('cmd.textToColumns'),
			(tx) => {
				textToColumns(tx, store.workbook, sheetIndex, area, options);
			},
			{ selectionAfter: after },
		);
	} catch (e) {
		if (e instanceof TextToColumnsError) {
			return e.message;
		}
		throw e;
	}
	return null;
}

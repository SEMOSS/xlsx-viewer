import type { Cell, CellAddress, RangeAddress, Worksheet } from 'xlsx-viewer-core';
import {
	cellKey,
	columnToLetters,
	formatValue,
	guessHeader,
	isCellError,
	isDateFormat,
	SortError,
	sortRange,
	usedRange,
} from 'xlsx-viewer-core';

import type { Translate } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';
import { dataBlock } from './structure-commands';

/** Excel's Sort dialog accepts up to 64 levels. */
export const MAX_SORT_LEVELS = 64;

/** How the Order box of a level is worded. */
export type SortValueKind = 'text' | 'number' | 'date';

/** One "Sort by" / "Then by" row of the Sort dialog. */
export interface SortLevel {
	/** Absolute column (row when sorting left to right); `null` until one is chosen. */
	key: number | null;
	descending: boolean;
}

/** A column (or row) the block can be sorted by. */
export interface SortKeyOption {
	index: number;
	label: string;
	kind: SortValueKind;
}

export interface SortDialogSetup {
	sheetIndex: number;
	/** The block to sort, or `null` when there is nothing to sort (see `error`). */
	block: RangeAddress | null;
	/** Why the dialog cannot be used on the selection. */
	error?: string;
	/** "My data has headers" as guessed from the data. */
	hasHeader: boolean;
	active: CellAddress;
}

export interface SortDialogOptions {
	hasHeader: boolean;
	byColumns: boolean;
	caseSensitive: boolean;
}

/** Stop scanning for a column's data type after this many cells. */
const KIND_SCAN_LIMIT = 10_000;

function isBlankCell(cell: Cell | undefined): boolean {
	return !cell || ((cell.value === null || cell.value === '') && !cell.formula);
}

/** The block without the empty rows / columns beyond the sheet's used range. */
function clampToUsed(sheet: Worksheet, block: RangeAddress): RangeAddress | null {
	const used = usedRange(sheet, true);
	if (!used) {
		return null;
	}
	const bottom = Math.min(block.bottom, used.bottom);
	const right = Math.min(block.right, used.right);
	return bottom < block.top || right < block.left ? null : { ...block, bottom, right };
}

/**
 * What the Sort dialog works on: the selection, or the current region of a
 * single selected cell, clamped to the used range.
 */
export function sortDialogSetup(store: ViewerStore): SortDialogSetup {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const sel = store.selectionOf();
	const base = { sheetIndex, hasHeader: false, active: sel.active };
	if (sel.ranges.length > 1) {
		return { ...base, block: null, error: store.t('dialog.sort.multipleSelections') };
	}
	const block = clampToUsed(sheet, dataBlock(store));
	const single = block && block.top === block.bottom && block.left === block.right;
	if (!block || (single && isBlankCell(sheet.cells.get(cellKey(block.top, block.left))))) {
		return { ...base, block: null, error: store.t('dialog.sort.nothingToSort') };
	}
	return { ...base, block, hasHeader: guessHeader(sheet, block) };
}

/** The part of the block that moves (without the header row). */
export function sortDataRange(
	block: RangeAddress,
	opts: Pick<SortDialogOptions, 'hasHeader' | 'byColumns'>,
): RangeAddress {
	if (!opts.hasHeader) {
		return block;
	}
	if (opts.byColumns) {
		return block.right > block.left ? { ...block, left: block.left + 1 } : block;
	}
	return block.bottom > block.top ? { ...block, top: block.top + 1 } : block;
}

/** Select what will be sorted, as Excel does while its Sort dialog is open. */
export function selectSortBlock(
	store: ViewerStore,
	block: RangeAddress,
	opts: Pick<SortDialogOptions, 'hasHeader' | 'byColumns'>,
): void {
	const data = sortDataRange(block, opts);
	const sel = store.selectionOf();
	const current = sel.ranges[sel.ranges.length - 1];
	const active = {
		row: Math.min(Math.max(sel.active.row, data.top), data.bottom),
		col: Math.min(Math.max(sel.active.col, data.left), data.right),
	};
	const same =
		sel.ranges.length === 1 &&
		current.top === data.top &&
		current.left === data.left &&
		current.bottom === data.bottom &&
		current.right === data.right &&
		active.row === sel.active.row &&
		active.col === sel.active.col;
	if (!same) {
		store.setSelection({ active, anchor: { row: data.top, col: data.left }, ranges: [data] });
	}
}

function cellText(store: ViewerStore, cell: Cell | undefined): string {
	const v = cell?.value ?? null;
	if (v === null) {
		return '';
	}
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (isCellError(v)) {
		return v.code;
	}
	const numFmt = store.styles.resolve(cell?.style).numFmt;
	return formatValue(v, numFmt, { date1904: store.workbook.date1904 }).text.trim();
}

/** Numbers, dates (numbers with a date format) or text, from the first data cell with a value. */
function valueKind(
	store: ViewerStore,
	cells: (i: number) => Cell | undefined,
	from: number,
	to: number,
): SortValueKind {
	const end = Math.min(to, from + KIND_SCAN_LIMIT);
	for (let i = from; i <= end; i++) {
		const cell = cells(i);
		const v = cell?.value;
		if (v === undefined || v === null || v === '') {
			continue;
		}
		if (typeof v !== 'number') {
			return 'text';
		}
		return isDateFormat(store.styles.resolve(cell?.style).numFmt) ? 'date' : 'number';
	}
	return 'text';
}

/**
 * The columns (rows when sorting left to right) of the block for the
 * level boxes: header text with "My data has headers", otherwise
 * "Column B" / "Row 3", each with the kind of values it holds.
 */
export function sortKeyOptions(
	store: ViewerStore,
	block: RangeAddress,
	opts: Pick<SortDialogOptions, 'hasHeader' | 'byColumns'>,
): SortKeyOption[] {
	const sheet = store.activeSheet;
	const data = sortDataRange(block, opts);
	const out: SortKeyOption[] = [];
	if (opts.byColumns) {
		for (let row = block.top; row <= block.bottom; row++) {
			const kind = valueKind(store, (c) => sheet.cells.get(cellKey(row, c)), data.left, data.right);
			out.push({ index: row, label: store.t('dialog.sort.rowNumber', { n: row + 1 }), kind });
		}
		return out;
	}
	const header = opts.hasHeader && data.top > block.top;
	for (let col = block.left; col <= block.right; col++) {
		const label = header ? cellText(store, sheet.cells.get(cellKey(block.top, col))) : '';
		const kind = valueKind(store, (r) => sheet.cells.get(cellKey(r, col)), data.top, data.bottom);
		const letter = columnToLetters(col);
		out.push({ index: col, label: label || store.t('dialog.sort.columnLetter', { letter }), kind });
	}
	return out;
}

/** Labels of the ascending / descending choices of the Order box. */
export function sortOrderLabels(kind: SortValueKind, t: Translate): [string, string] {
	switch (kind) {
		case 'number':
			return [t('dialog.sort.smallestToLargest'), t('dialog.sort.largestToSmallest')];
		case 'date':
			return [t('dialog.sort.oldestToNewest'), t('dialog.sort.newestToOldest')];
		default:
			return [t('dialog.sort.aToZ'), t('dialog.sort.zToA')];
	}
}

/** Excel's complaint about the levels, or `null` when they can be applied. */
export function checkSortLevels(
	levels: readonly SortLevel[],
	options: readonly SortKeyOption[],
	t: Translate,
): string | null {
	if (levels.some((level) => level.key === null)) {
		return t('dialog.sort.columnRequired');
	}
	const seen = new Set<number>();
	for (const level of levels) {
		const key = level.key as number;
		if (seen.has(key)) {
			const label =
				options.find((o) => o.index === key)?.label ??
				t('dialog.sort.columnLetter', { letter: columnToLetters(key) });
			return t('dialog.sort.duplicateCondition', { label });
		}
		seen.add(key);
	}
	return null;
}

/**
 * OK in the Sort dialog: sort the block by the levels in one undoable step.
 * Returns the message to show instead (bad levels, merged cells), or `null`.
 */
export function applySortLevels(
	store: ViewerStore,
	block: RangeAddress,
	levels: readonly SortLevel[],
	opts: SortDialogOptions,
	options: readonly SortKeyOption[],
): string | null {
	const error = checkSortLevels(levels, options, store.t);
	if (error) {
		return error;
	}
	if (levels.length === 0) {
		return null;
	}
	const sheetIndex = store.state.activeSheet;
	const keys = levels.map((level) => ({
		index: level.key as number,
		descending: level.descending,
	}));
	try {
		store.transact(store.t('menu.sort'), (tx) => {
			sortRange(tx, store.workbook, sheetIndex, block, keys, {
				hasHeader: opts.hasHeader,
				byColumns: opts.byColumns,
				caseSensitive: opts.caseSensitive,
			});
		});
	} catch (e) {
		if (e instanceof SortError) {
			return e.message;
		}
		throw e;
	}
	return null;
}

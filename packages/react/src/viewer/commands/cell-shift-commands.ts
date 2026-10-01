import type { CellShiftDirection, RangeAddress, Workbook } from 'xlsx-viewer-core';
import {
	CellShiftError,
	checkCellShift,
	deleteCellsShift,
	EditTransaction,
	insertCellsShift,
	isWholeColumnRange,
	isWholeRowRange,
	rangesIntersect,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';
import { deleteCellsCmd, insertCellsCmd } from './structure-commands';

type Apply = (tx: EditTransaction, workbook: Workbook, range: RangeAddress) => void;

/** Bottom-most (right-most) ranges first, so earlier shifts do not move later ranges. */
function ordered(ranges: readonly RangeAddress[], vertical: boolean): RangeAddress[] {
	return [...ranges].sort((a, b) =>
		vertical ? b.top - a.top || b.left - a.left : b.left - a.left || b.top - a.top,
	);
}

function overlapping(ranges: readonly RangeAddress[]): boolean {
	for (let i = 0; i < ranges.length; i++) {
		for (let j = i + 1; j < ranges.length; j++) {
			if (rangesIntersect(ranges[i], ranges[j])) {
				return true;
			}
		}
	}
	return false;
}

/**
 * A copy of the workbook that cell shifts can run on without touching the
 * original: shifts replace cells and collections, they never mutate shared
 * objects in place.
 */
function shadowOf(wb: Workbook): Workbook {
	return {
		...wb,
		sheets: wb.sheets.map((s) => ({ ...s, cells: new Map(s.cells), rows: new Map(s.rows) })),
		definedNames: [...wb.definedNames],
	};
}

/**
 * Run every range on a shadow copy to find a refusal that only shows up
 * after an earlier range has shifted (the real workbook is not touched).
 */
function trialRun(
	store: ViewerStore,
	ranges: readonly RangeAddress[],
	apply: Apply,
): CellShiftError | null {
	const shadow = shadowOf(store.workbook);
	const tx = new EditTransaction(shadow, 'trial');
	try {
		for (const range of ranges) {
			apply(tx, shadow, range);
		}
	} catch (e) {
		if (e instanceof CellShiftError) {
			return e;
		}
		throw e;
	}
	return null;
}

/**
 * Entire rows / columns selected: Excel inserts or deletes whole rows /
 * columns whatever the shift direction.
 */
export function wholeAxisOfRanges(ranges: readonly RangeAddress[]): 'rows' | 'cols' | null {
	if (ranges.length > 0 && ranges.every(isWholeRowRange)) {
		return 'rows';
	}
	return ranges.length > 0 && ranges.every(isWholeColumnRange) ? 'cols' : null;
}

function runShift(
	store: ViewerStore,
	label: string,
	shift: CellShiftDirection,
	apply: Apply,
): boolean {
	if (store.readOnly) {
		return false;
	}
	const sheetIndex = store.state.activeSheet;
	const ranges = ordered(store.selectionOf().ranges, shift === 'down' || shift === 'up');
	if (overlapping(ranges)) {
		// Excel's message when the selected areas overlap.
		store.setUi({ message: store.t('message.overlappingSelections') });
		return false;
	}
	let problem: CellShiftError | null = null;
	for (const range of ranges) {
		problem ??= checkCellShift(store.workbook, sheetIndex, range, shift);
	}
	if (!problem && ranges.length > 1) {
		problem = trialRun(store, ranges, apply);
	}
	if (problem) {
		// Excel's refusal, worded by the core (CELL_SHIFT_MESSAGES).
		store.setUi({ message: problem.message });
		return false;
	}
	return store.transact(label, (tx) => {
		for (const range of ranges) {
			apply(tx, store.workbook, range);
		}
	});
}

/**
 * Insert Cells with "Shift cells right" / "Shift cells down" at every
 * selected range, as one undoable step (entire rows / columns when whole
 * rows / columns are selected). Shows Excel's message and changes nothing
 * when the shift is refused. Returns true when applied.
 */
export function insertCellsShiftCmd(store: ViewerStore, shift: 'right' | 'down'): boolean {
	const sheetIndex = store.state.activeSheet;
	const axis = wholeAxisOfRanges(store.selectionOf().ranges);
	if (axis && !store.readOnly) {
		insertCellsCmd(store, axis);
		return true;
	}
	return runShift(store, store.t('undo.insertCells'), shift, (tx, wb, range) =>
		insertCellsShift(tx, wb, sheetIndex, range, shift),
	);
}

/**
 * Delete Cells with "Shift cells left" / "Shift cells up" at every selected
 * range, as one undoable step (entire rows / columns when whole rows /
 * columns are selected). Shows Excel's message and changes nothing when the
 * shift is refused. Returns true when applied.
 */
export function deleteCellsShiftCmd(store: ViewerStore, shift: 'left' | 'up'): boolean {
	const sheetIndex = store.state.activeSheet;
	const axis = wholeAxisOfRanges(store.selectionOf().ranges);
	if (axis && !store.readOnly) {
		deleteCellsCmd(store, axis);
		return true;
	}
	return runShift(store, store.t('undo.deleteCells'), shift, (tx, wb, range) =>
		deleteCellsShift(tx, wb, sheetIndex, range, shift),
	);
}

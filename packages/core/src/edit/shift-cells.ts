import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { cellRange, isSingleCell, rangeContains } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import {
	adjustRangeForBand,
	bandFateOf,
	bandShiftOf,
	bandZones,
	type BandZones,
	type CellBandShift,
	type CellShiftDirection,
	movePointForBand,
} from './shift-band';
import { adjustHomeSheetForBand, adjustWorkbookForBand, tableFateForBand } from './shift-features';
import type { EditTransaction } from './transaction';

export type { CellShiftDirection } from './shift-band';

/** Why Excel refuses to insert / delete cells with shifting. */
export type CellShiftReason = 'merge' | 'dataLoss' | 'array' | 'table';

/** Excel's messages for each refusal. */
export const CELL_SHIFT_MESSAGES: Readonly<Record<CellShiftReason, string>> = {
	merge: "We can't do that to a merged cell.",
	dataLoss:
		'To prevent possible loss of data, Excel cannot shift nonblank cells off of the worksheet. ' +
		'Select another location in which to insert new cells, or delete data from the end of your ' +
		'worksheet.',
	array: "You can't change part of an array.",
	table:
		'This operation is not allowed. The operation is attempting to shift cells in a table on ' +
		'your worksheet.',
};

export class CellShiftError extends Error {
	constructor(
		message: string,
		readonly reason: CellShiftReason,
	) {
		super(message);
		this.name = 'CellShiftError';
	}
}

function shiftError(reason: CellShiftReason): CellShiftError {
	return new CellShiftError(CELL_SHIFT_MESSAGES[reason], reason);
}

function isNonBlank(cell: Cell): boolean {
	return cell.formula !== undefined || (cell.value !== null && cell.value !== '');
}

function clampToGrid(r: RangeAddress): RangeAddress {
	return {
		top: Math.max(0, r.top),
		left: Math.max(0, r.left),
		bottom: Math.min(MAX_ROW_INDEX, r.bottom),
		right: Math.min(MAX_COL_INDEX, r.right),
	};
}

/**
 * Check whether Insert / Delete cells with `shift` at `range` is allowed,
 * without changing anything. Returns the refusal Excel would show, or
 * `null`: merged cells, legacy array formulas or tables that would be split,
 * and (insertions) nonblank cells that would be pushed off the grid.
 */
export function checkCellShift(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	shift: CellShiftDirection,
): CellShiftError | null {
	const sheet = workbook.sheets[sheetIndex];
	const band = bandShiftOf(sheet.name, clampToGrid(range), shift);
	const zones = bandZones(band);
	if (sheet.merges.some((m) => bandFateOf(m, zones) === 'split')) {
		return shiftError('merge');
	}
	const limit = (band.vertical ? MAX_ROW_INDEX : MAX_COL_INDEX) - band.count;
	let arraySplit = false;
	let dataLoss = false;
	for (const [key, cell] of sheet.cells) {
		const ref = cell.array?.ref;
		if (ref && !cell.array?.dynamic && !isSingleCell(ref) && bandFateOf(ref, zones) === 'split') {
			arraySplit = true;
		}
		if (band.insert && !dataLoss && isNonBlank(cell)) {
			const line = band.vertical ? keyRow(key) : keyCol(key);
			const cross = band.vertical ? keyCol(key) : keyRow(key);
			dataLoss = cross >= band.from && cross <= band.to && line >= band.at && line > limit;
		}
	}
	if (arraySplit) {
		return shiftError('array');
	}
	if (sheet.tables.some((t) => tableFateForBand(t, band, zones).kind === 'split')) {
		return shiftError('table');
	}
	return dataLoss ? shiftError('dataLoss') : null;
}

function forEachCellIn(
	cells: Map<number, Cell>,
	range: RangeAddress,
	visit: (key: number, cell: Cell) => void,
): void {
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	if (area > cells.size) {
		for (const [key, cell] of [...cells]) {
			if (rangeContains(range, keyRow(key), keyCol(key))) {
				visit(key, cell);
			}
		}
		return;
	}
	for (let row = range.top; row <= range.bottom; row++) {
		for (let col = range.left; col <= range.right; col++) {
			const key = cellKey(row, col);
			const cell = cells.get(key);
			if (cell) {
				visit(key, cell);
			}
		}
	}
}

/** A spilled value without its formula: keep only the format. */
function withoutSpill(cell: Cell): Cell | undefined {
	return cell.style ? { value: null, style: cell.style } : undefined;
}

function putCell(cells: Map<number, Cell>, key: number, cell: Cell | undefined): void {
	if (cell) {
		cells.set(key, cell);
	} else {
		cells.delete(key);
	}
}

/**
 * Dynamic arrays whose spill range would be split lose their spilled values
 * and spill again from their (moved) formula cell on recalculation.
 */
function collapseSplitSpills(sheet: Worksheet, zones: BandZones): void {
	const anchors: number[] = [];
	for (const [key, cell] of sheet.cells) {
		const ref = cell.array?.ref;
		if (cell.array?.dynamic && ref && !isSingleCell(ref) && bandFateOf(ref, zones) === 'split') {
			anchors.push(key);
		}
	}
	for (const key of anchors) {
		const anchor = sheet.cells.get(key);
		if (!anchor?.array) {
			continue;
		}
		forEachCellIn(sheet.cells, anchor.array.ref, (k, member) => {
			if (member.arrayAnchor === key) {
				putCell(sheet.cells, k, withoutSpill(member));
			}
		});
		const ref = cellRange(keyRow(key), keyCol(key));
		sheet.cells.set(key, { ...anchor, array: { ...anchor.array, ref } });
	}
}

/** Move the cells of the band (deleting removed ones) and fix array bookkeeping. */
function moveBandCells(sheet: Worksheet, band: CellBandShift): void {
	const next = new Map<number, Cell>();
	const remap = new Map<number, number>();
	const removed = new Set<number>();
	for (const [key, cell] of sheet.cells) {
		const at = { row: keyRow(key), col: keyCol(key) };
		const moved = movePointForBand(at, band);
		if (!moved) {
			removed.add(key);
			continue;
		}
		const nextKey = moved === at ? key : cellKey(moved.row, moved.col);
		if (nextKey !== key) {
			remap.set(key, nextKey);
		}
		next.set(nextKey, cell);
	}
	for (const [key, cell] of next) {
		let fixed: Cell | undefined = cell;
		if (cell.arrayAnchor !== undefined) {
			const anchor = remap.get(cell.arrayAnchor);
			if (removed.has(cell.arrayAnchor)) {
				fixed = withoutSpill(cell);
			} else if (anchor !== undefined) {
				fixed = { ...cell, arrayAnchor: anchor };
			}
		}
		if (fixed?.array) {
			const ref = adjustRangeForBand(fixed.array.ref, band);
			if (ref && ref !== fixed.array.ref) {
				fixed = { ...fixed, array: { ...fixed.array, ref } };
			}
		}
		if (fixed !== cell) {
			putCell(next, key, fixed);
		}
	}
	sheet.cells = next;
}

/** Inserted cells take the format of the cells above (left of) them, like Excel's default. */
function copyFormatsIntoInserted(sheet: Worksheet, band: CellBandShift): void {
	if (band.at === 0) {
		return;
	}
	const source = band.at - 1;
	const last = Math.min(band.at + band.count - 1, band.vertical ? MAX_ROW_INDEX : MAX_COL_INDEX);
	const keyOf = (line: number, cross: number) =>
		band.vertical ? cellKey(line, cross) : cellKey(cross, line);
	const styled: Array<[cross: number, style: number]> = [];
	if (band.to - band.from + 1 <= sheet.cells.size) {
		for (let cross = band.from; cross <= band.to; cross++) {
			const style = sheet.cells.get(keyOf(source, cross))?.style;
			if (style) {
				styled.push([cross, style]);
			}
		}
	} else {
		for (const [key, cell] of sheet.cells) {
			const line = band.vertical ? keyRow(key) : keyCol(key);
			const cross = band.vertical ? keyCol(key) : keyRow(key);
			if (line === source && cross >= band.from && cross <= band.to && cell.style) {
				styled.push([cross, cell.style]);
			}
		}
	}
	for (const [cross, style] of styled) {
		for (let line = band.at; line <= last; line++) {
			sheet.cells.set(keyOf(line, cross), { value: null, style });
		}
	}
}

function shiftCells(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	shift: CellShiftDirection,
): void {
	const problem = checkCellShift(workbook, sheetIndex, range, shift);
	if (problem) {
		throw problem;
	}
	const sheet = workbook.sheets[sheetIndex];
	const band = bandShiftOf(sheet.name, clampToGrid(range), shift);
	const zones = bandZones(band);
	for (let i = 0; i < workbook.sheets.length; i++) {
		tx.touchSheet(i, true);
	}
	tx.touchWorkbook();
	collapseSplitSpills(sheet, zones);
	moveBandCells(sheet, band);
	if (band.insert) {
		copyFormatsIntoInserted(sheet, band);
	}
	adjustWorkbookForBand(workbook, band, sheet);
	adjustHomeSheetForBand(sheet, band, zones);
	sheet.dirty = true;
}

/**
 * Home > Insert > Insert Cells with "Shift cells right" / "Shift cells
 * down": blank cells appear at `range` and the cells of its rows (columns)
 * move right (down). References across the workbook follow the moved cells.
 * Throws `CellShiftError` (before changing anything) when Excel refuses.
 */
export function insertCellsShift(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	shift: 'right' | 'down',
): void {
	shiftCells(tx, workbook, sheetIndex, range, shift);
}

/**
 * Home > Delete > Delete Cells with "Shift cells left" / "Shift cells up":
 * the cells of `range` are removed and the cells right of (below) them move
 * in. References to deleted cells become `#REF!`. Throws `CellShiftError`
 * (before changing anything) when Excel refuses.
 */
export function deleteCellsShift(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	shift: 'left' | 'up',
): void {
	shiftCells(tx, workbook, sheetIndex, range, shift);
}

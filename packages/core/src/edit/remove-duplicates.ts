import { cellKey, keyCol, keyRow } from '../address/cell-address';
import {
	isSingleCell,
	isWholeColumnRange,
	isWholeRowRange,
	rangeContains,
	rangesIntersect,
} from '../address/range';
import { translateFormula } from '../formula/translate';
import { formatValue } from '../numfmt/format-value';
import { usedRange } from '../sheet/worksheet';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

export interface RemoveDuplicatesOptions {
	/** Absolute indices of the compared columns (columns outside the range are ignored). */
	columns: number[];
	/** The first row of the range holds headers: it is neither compared nor moved. */
	hasHeader: boolean;
}

export interface RemoveDuplicatesResult {
	/** Duplicate rows removed. */
	removed: number;
	/** Unique data rows left (the header row is not counted). */
	remaining: number;
}

export interface DuplicateRows {
	/** Data rows kept (absolute row indices, in order): the first occurrence of each value. */
	kept: number[];
	/** Data rows repeating an earlier row in every compared column. */
	duplicates: number[];
	/** The range actually examined (trimmed to the used part of the sheet). */
	range: RangeAddress;
}

export class RemoveDuplicatesError extends Error {
	constructor(
		message: string,
		readonly reason: 'merge' | 'array',
	) {
		super(message);
		this.name = 'RemoveDuplicatesError';
	}
}

/** Whole-column or whole-row selections only cover the used part of the sheet, like Excel. */
export function trimRangeToUsed(sheet: Worksheet, range: RangeAddress): RangeAddress {
	const rows = isWholeColumnRange(range);
	const cols = isWholeRowRange(range);
	if (!rows && !cols) {
		return range;
	}
	const used = usedRange(sheet) ?? { top: 0, left: 0, bottom: 0, right: 0 };
	return {
		...range,
		bottom: rows ? Math.max(range.top, Math.min(range.bottom, used.bottom)) : range.bottom,
		right: cols ? Math.max(range.left, Math.min(range.right, used.right)) : range.right,
	};
}

/** Part of a multi-cell array formula (legacy or spilled), which cannot move row by row. */
function inMultiCellArray(cell: Cell): boolean {
	return (
		cell.arrayAnchor !== undefined || (cell.array !== undefined && !isSingleCell(cell.array.ref))
	);
}

function validate(sheet: Worksheet, range: RangeAddress): void {
	if (sheet.merges.some((m) => rangesIntersect(m, range))) {
		throw new RemoveDuplicatesError("We can't do that to a merged cell.", 'merge');
	}
	for (const [key, cell] of sheet.cells) {
		if (inMultiCellArray(cell) && rangeContains(range, keyRow(key), keyCol(key))) {
			throw new RemoveDuplicatesError("You can't change part of an array.", 'array');
		}
	}
}

/** The text Excel compares: the displayed value, case-insensitively. */
function displayKey(workbook: Workbook, cell: Cell | undefined): string {
	if (!cell || cell.value === null) {
		return '';
	}
	const xf = workbook.styles.cellXfs[cell.style ?? 0];
	const code = numberFormatCode(xf?.numFmtId ?? 0, workbook.styles.numFmts);
	return formatValue(cell.value, code, { date1904: workbook.date1904 }).text.toLowerCase();
}

/**
 * Find the rows of `range` that repeat an earlier row in every compared
 * column, comparing displayed text case-insensitively like Excel's Data >
 * Remove Duplicates. Throws `RemoveDuplicatesError` for merged cells or
 * multi-cell array formulas in the range.
 */
export function findDuplicateRows(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	options: RemoveDuplicatesOptions,
): DuplicateRows {
	const sheet = workbook.sheets[sheetIndex];
	const area = trimRangeToUsed(sheet, range);
	validate(sheet, area);
	const columns = [...new Set(options.columns)].filter((c) => c >= area.left && c <= area.right);
	const seen = new Set<string>();
	const kept: number[] = [];
	const duplicates: number[] = [];
	for (let row = area.top + (options.hasHeader ? 1 : 0); row <= area.bottom; row++) {
		const key = columns
			.map((c) => displayKey(workbook, sheet.cells.get(cellKey(row, c))))
			.join('\u0000');
		if (columns.length > 0 && seen.has(key)) {
			duplicates.push(row);
		} else {
			seen.add(key);
			kept.push(row);
		}
	}
	return { kept, duplicates, range: area };
}

/** A cell copied `dRow` rows away: formulas are translated as if copied, like sorting. */
function movedCell(cell: Cell, dRow: number): Cell {
	if (cell.formula === undefined) {
		return cell;
	}
	const next: Cell = { ...cell, formula: translateFormula(cell.formula, dRow, 0) };
	if (cell.array) {
		const ref = cell.array.ref;
		next.array = { ...cell.array, ref: { ...ref, top: ref.top + dRow, bottom: ref.bottom + dRow } };
	}
	return next;
}

/**
 * Data > Remove Duplicates: delete the rows of `range` that repeat an
 * earlier row in the compared columns. Kept rows move up inside the range
 * (values, formats and formulas, translated as if copied); the rows freed
 * at the bottom of the range are emptied. Cells outside the range are not
 * touched. Throws `RemoveDuplicatesError` before changing anything.
 */
export function removeDuplicateRows(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	options: RemoveDuplicatesOptions,
): RemoveDuplicatesResult {
	const found = findDuplicateRows(workbook, sheetIndex, range, options);
	const { kept, duplicates, range: area } = found;
	if (duplicates.length === 0) {
		return { removed: 0, remaining: kept.length };
	}
	const sheet = workbook.sheets[sheetIndex];
	const first = area.top + (options.hasHeader ? 1 : 0);
	tx.touchRange(sheetIndex, { ...area, top: first });
	const width = area.right - area.left + 1;
	const rows = kept.map((row) => {
		const cells: Array<Cell | undefined> = [];
		for (let i = 0; i < width; i++) {
			cells.push(sheet.cells.get(cellKey(row, area.left + i)));
		}
		return { row, cells };
	});
	for (let target = first; target <= area.bottom; target++) {
		const source = rows[target - first];
		for (let i = 0; i < width; i++) {
			const key = cellKey(target, area.left + i);
			const cell = source?.cells[i];
			if (!cell) {
				sheet.cells.delete(key);
			} else if (source.row !== target) {
				sheet.cells.set(key, movedCell(cell, target - source.row));
			}
		}
	}
	sheet.dirty = true;
	return { removed: duplicates.length, remaining: kept.length };
}

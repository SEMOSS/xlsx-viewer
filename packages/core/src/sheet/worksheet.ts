import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX } from '../address/constants';
import type { RangeAddress } from '../types/address';
import type { Cell, CellValue } from '../types/cell';
import type { ColInfo, RowInfo, SheetView, Worksheet } from '../types/sheet';

let idCounter = 0;

/** A process-unique id for in-memory objects (sheets, drawing objects). */
export function generateId(prefix = 'id'): string {
	idCounter += 1;
	const rand = Math.random().toString(36).slice(2, 8);
	return `${prefix}-${Date.now().toString(36)}-${idCounter}-${rand}`;
}

export function createDefaultSheetView(): SheetView {
	return {
		showGridLines: true,
		showRowColHeaders: true,
		showZeros: true,
		showFormulas: false,
		rightToLeft: false,
		zoomScale: 100,
	};
}

/** A new empty worksheet. */
export function createWorksheet(name: string, sheetId: number): Worksheet {
	return {
		id: generateId('sheet'),
		name,
		sheetId,
		kind: 'worksheet',
		state: 'visible',
		cells: new Map(),
		rows: new Map(),
		cols: [],
		defaultRowHeight: 15,
		merges: [],
		view: createDefaultSheetView(),
		hyperlinks: [],
		conditionalFormats: [],
		dataValidations: [],
		tables: [],
		drawings: [],
		notes: [],
		threadedComments: [],
		sparklineGroups: [],
		dirty: true,
	};
}

export function getCell(sheet: Worksheet, row: number, col: number): Cell | undefined {
	return sheet.cells.get(cellKey(row, col));
}

export function getCellValue(sheet: Worksheet, row: number, col: number): CellValue {
	return sheet.cells.get(cellKey(row, col))?.value ?? null;
}

/** Store a cell (or delete it when `cell` is undefined). Marks the sheet dirty. */
export function putCell(sheet: Worksheet, row: number, col: number, cell: Cell | undefined): void {
	const key = cellKey(row, col);
	if (cell === undefined) {
		sheet.cells.delete(key);
	} else {
		sheet.cells.set(key, cell);
	}
	sheet.dirty = true;
}

/** A cell is "empty" when it has no value, no formula and no explicit style. */
export function isCellEmpty(cell: Cell | undefined): boolean {
	return !cell || (cell.value === null && !cell.formula && !cell.style && !cell.richText);
}

/** Cell keys sorted in row-major order. */
export function sortedCellKeys(sheet: Worksheet): number[] {
	return [...sheet.cells.keys()].sort((a, b) => a - b);
}

/**
 * Bounding range of all cells that have content (value or formula), or of
 * any stored cell when `includeFormatted` is set. Returns `null` for an
 * empty sheet.
 */
export function usedRange(sheet: Worksheet, includeFormatted = false): RangeAddress | null {
	let top = Infinity;
	let left = Infinity;
	let bottom = -1;
	let right = -1;
	for (const [key, cell] of sheet.cells) {
		if (!includeFormatted && cell.value === null && !cell.formula) {
			continue;
		}
		const r = keyRow(key);
		const c = keyCol(key);
		if (r < top) {
			top = r;
		}
		if (r > bottom) {
			bottom = r;
		}
		if (c < left) {
			left = c;
		}
		if (c > right) {
			right = c;
		}
	}
	if (bottom < 0) {
		return null;
	}
	return { top, left, bottom, right };
}

/** Column record covering `col`, if any (binary search over sorted records). */
export function getColInfo(sheet: Worksheet, col: number): ColInfo | undefined {
	const cols = sheet.cols;
	let lo = 0;
	let hi = cols.length - 1;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		const c = cols[mid];
		if (col < c.min) {
			hi = mid - 1;
		} else if (col > c.max) {
			lo = mid + 1;
		} else {
			return c;
		}
	}
	return undefined;
}

export function getRowInfo(sheet: Worksheet, row: number): RowInfo | undefined {
	return sheet.rows.get(row);
}

/** Split column records so that `[min, max]` is covered by records aligned to its edges. */
function splitColsAt(sheet: Worksheet, min: number, max: number): void {
	const out: ColInfo[] = [];
	for (const c of sheet.cols) {
		if (c.max < min || c.min > max) {
			out.push(c);
			continue;
		}
		if (c.min < min) {
			out.push({ ...c, max: min - 1 });
		}
		out.push({ ...c, min: Math.max(c.min, min), max: Math.min(c.max, max) });
		if (c.max > max) {
			out.push({ ...c, min: max + 1 });
		}
	}
	sheet.cols = out.sort((a, b) => a.min - b.min);
}

/**
 * Apply a change to the column records for columns `min..max`, creating
 * records for uncovered columns. Adjacent identical records are merged.
 */
export function updateCols(
	sheet: Worksheet,
	min: number,
	max: number,
	update: (info: ColInfo) => ColInfo,
): void {
	const lo = Math.max(0, min);
	const hi = Math.min(MAX_COL_INDEX, max);
	splitColsAt(sheet, lo, hi);
	const result: ColInfo[] = [];
	let cursor = lo;
	for (const c of sheet.cols) {
		if (c.max < lo || c.min > hi) {
			result.push(c);
			continue;
		}
		if (c.min > cursor) {
			result.push(update({ min: cursor, max: c.min - 1 }));
		}
		result.push(update({ ...c }));
		cursor = c.max + 1;
	}
	if (cursor <= hi) {
		result.push(update({ min: cursor, max: hi }));
	}
	result.sort((a, b) => a.min - b.min);
	sheet.cols = mergeAdjacentCols(result);
	sheet.dirty = true;
}

function colKey(c: ColInfo): string {
	return JSON.stringify([
		c.width,
		c.customWidth,
		c.hidden,
		c.style,
		c.outlineLevel,
		c.collapsed,
		c.bestFit,
	]);
}

function mergeAdjacentCols(cols: ColInfo[]): ColInfo[] {
	const out: ColInfo[] = [];
	for (const c of cols) {
		const isEmpty =
			c.width === undefined &&
			!c.hidden &&
			c.style === undefined &&
			!c.outlineLevel &&
			!c.collapsed &&
			!c.customWidth;
		if (isEmpty) {
			continue;
		}
		const prev = out[out.length - 1];
		if (prev && prev.max + 1 === c.min && colKey(prev) === colKey(c)) {
			prev.max = c.max;
		} else {
			out.push({ ...c });
		}
	}
	return out;
}

/** Update a row record, creating it when needed and deleting it when it becomes empty. */
export function updateRow(sheet: Worksheet, row: number, update: (info: RowInfo) => RowInfo): void {
	const next = update({ ...(sheet.rows.get(row) ?? {}) });
	const isEmpty = Object.values(next).every((v) => v === undefined || v === false);
	if (isEmpty) {
		sheet.rows.delete(row);
	} else {
		sheet.rows.set(row, next);
	}
	sheet.dirty = true;
}

/** Merged range containing a cell, if any. */
export function findMerge(sheet: Worksheet, row: number, col: number): RangeAddress | undefined {
	for (const m of sheet.merges) {
		if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
			return m;
		}
	}
	return undefined;
}

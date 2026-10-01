import { cellKey } from '../address/cell-address';
import { rangesIntersect } from '../address/range';
import { compareText } from '../formula/coerce';
import { translateFormula } from '../formula/translate';
import type { RangeAddress } from '../types/address';
import type { Cell, CellValue } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import type { EditTransaction } from './transaction';

export interface SortKey {
	/** Absolute column index (row index when `byColumns`). */
	index: number;
	descending?: boolean;
}

export interface SortOptions {
	/** The first row (column) holds headers and stays in place. */
	hasHeader?: boolean;
	/** Sort left to right (columns) instead of top to bottom. */
	byColumns?: boolean;
	caseSensitive?: boolean;
}

export class SortError extends Error {
	constructor(
		message: string,
		readonly reason: 'merged',
	) {
		super(message);
		this.name = 'SortError';
	}
}

function rank(v: CellValue): number {
	if (v === null || v === '') {
		return 4;
	}
	if (typeof v === 'number') {
		return 0;
	}
	if (typeof v === 'string') {
		return 1;
	}
	if (typeof v === 'boolean') {
		return 2;
	}
	return 3;
}

/**
 * Case-sensitive tie-break for texts equal ignoring case: Excel puts the
 * lowercase letter first at the first position where the case differs
 * ("apple" < "Apple" < "APPLE").
 */
function caseTieBreak(a: string, b: string): number {
	for (let i = 0; i < Math.min(a.length, b.length); i++) {
		const x = a[i];
		const y = b[i];
		if (x !== y) {
			const xLower = x === x.toLowerCase();
			const yLower = y === y.toLowerCase();
			if (xLower !== yLower) {
				return xLower ? -1 : 1;
			}
		}
	}
	return 0;
}

/** Excel's sort order: numbers < text < logicals < errors, blanks always last. */
export function compareForSort(
	a: CellValue,
	b: CellValue,
	descending: boolean,
	caseSensitive = false,
): number {
	const ra = rank(a);
	const rb = rank(b);
	if (ra === 4 || rb === 4) {
		return ra === rb ? 0 : ra === 4 ? 1 : -1;
	}
	let c: number;
	if (ra !== rb) {
		c = ra - rb;
	} else if (typeof a === 'number' && typeof b === 'number') {
		c = a - b;
	} else if (typeof a === 'string' && typeof b === 'string') {
		c = compareText(a, b) || (caseSensitive ? caseTieBreak(a, b) : 0);
	} else if (typeof a === 'boolean' && typeof b === 'boolean') {
		c = a === b ? 0 : a ? 1 : -1;
	} else if (isCellError(a) && isCellError(b)) {
		c = 0;
	} else {
		c = 0;
	}
	return descending ? -c : c;
}

/** Guess whether the first row of a range is a header row (text above differently typed data). */
export function guessHeader(sheet: Worksheet, range: RangeAddress): boolean {
	if (range.bottom <= range.top) {
		return false;
	}
	let textHeaders = 0;
	let dataNonText = 0;
	for (let c = range.left; c <= range.right; c++) {
		const h = sheet.cells.get(cellKey(range.top, c))?.value;
		const d = sheet.cells.get(cellKey(range.top + 1, c))?.value;
		if (typeof h === 'string' && h !== '') {
			textHeaders++;
		}
		if (d !== undefined && d !== null && typeof d !== 'string') {
			dataNonText++;
		}
	}
	const width = range.right - range.left + 1;
	return textHeaders === width && dataNonText > 0;
}

/**
 * Sort the rows (or columns) of a range by one or more keys, moving cell
 * contents and formats together. Formulas are adjusted as if copied to their
 * new row, which is what Excel does. Throws `SortError` for merged cells.
 */
export function sortRange(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	keys: SortKey[],
	opts: SortOptions = {},
): void {
	const sheet = workbook.sheets[sheetIndex];
	if (sheet.merges.some((m) => rangesIntersect(m, range))) {
		throw new SortError('To do this, all the merged cells need to be the same size.', 'merged');
	}
	const byCols = opts.byColumns === true;
	const data: RangeAddress = opts.hasHeader
		? byCols
			? { ...range, left: range.left + 1 }
			: { ...range, top: range.top + 1 }
		: range;
	const lines = byCols ? data.right - data.left + 1 : data.bottom - data.top + 1;
	const width = byCols ? data.bottom - data.top + 1 : data.right - data.left + 1;
	if (lines <= 1) {
		return;
	}
	tx.checkSheet(sheetIndex, 'sort');
	tx.touchRange(sheetIndex, data);
	const get = (line: number, i: number): Cell | undefined =>
		byCols
			? sheet.cells.get(cellKey(data.top + i, data.left + line))
			: sheet.cells.get(cellKey(data.top + line, data.left + i));
	const snapshot: Array<Array<Cell | undefined>> = [];
	for (let l = 0; l < lines; l++) {
		const row: Array<Cell | undefined> = [];
		for (let i = 0; i < width; i++) {
			row.push(get(l, i));
		}
		snapshot.push(row);
	}
	const keyValues = (l: number): CellValue[] =>
		keys.map((k) => {
			const i = byCols ? k.index - data.top : k.index - data.left;
			return snapshot[l][i]?.value ?? null;
		});
	const order = Array.from({ length: lines }, (_, i) => i);
	const values = order.map(keyValues);
	order.sort((a, b) => {
		for (let k = 0; k < keys.length; k++) {
			const c = compareForSort(
				values[a][k],
				values[b][k],
				keys[k].descending === true,
				opts.caseSensitive,
			);
			if (c !== 0) {
				return c;
			}
		}
		return a - b;
	});
	for (let target = 0; target < lines; target++) {
		const source = order[target];
		for (let i = 0; i < width; i++) {
			const cell = snapshot[source][i];
			const row = byCols ? data.top + i : data.top + target;
			const col = byCols ? data.left + target : data.left + i;
			const key = cellKey(row, col);
			if (!cell) {
				sheet.cells.delete(key);
				continue;
			}
			if (cell.formula && source !== target) {
				const dRow = byCols ? 0 : target - source;
				const dCol = byCols ? target - source : 0;
				sheet.cells.set(key, { ...cell, formula: translateFormula(cell.formula, dRow, dCol) });
			} else {
				sheet.cells.set(key, cell);
			}
		}
	}
	sheet.dirty = true;
}

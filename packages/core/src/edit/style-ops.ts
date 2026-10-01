import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { updateCols, updateRow } from '../sheet/worksheet';
import { StylePool, type StylePatch } from '../styles/style-pool';
import type { RangeAddress } from '../types/address';
import type { BorderEdge, BorderSpec } from '../types/style';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

function isWholeColumns(r: RangeAddress): boolean {
	return r.top === 0 && r.bottom === MAX_ROW_INDEX;
}

function isWholeRows(r: RangeAddress): boolean {
	return r.left === 0 && r.right === MAX_COL_INDEX;
}

/**
 * Re-style the cells of ranges with `derive(currentXfId) -> newXfId`. Whole
 * rows / columns update the row or column default style plus every existing
 * cell in them; other ranges create styled cells as needed. Results are
 * cached per base style.
 */
export function restyleRanges(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
	deriveStyle: (base: number | undefined) => number,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const cache = new Map<number, number>();
	const derive = (base: number | undefined): number => {
		const k = base ?? 0;
		const hit = cache.get(k);
		if (hit !== undefined) {
			return hit;
		}
		const id = deriveStyle(base);
		cache.set(k, id);
		return id;
	};
	for (const range of ranges) {
		if (isWholeColumns(range) || isWholeRows(range)) {
			tx.touchSheet(sheetIndex);
			applyToWholeLines(sheet, range, derive);
			continue;
		}
		tx.touchRange(sheetIndex, range, false);
		for (let row = range.top; row <= range.bottom; row++) {
			for (let col = range.left; col <= range.right; col++) {
				const key = cellKey(row, col);
				const cell = sheet.cells.get(key);
				const base = cell?.style ?? lineStyle(sheet, row, col);
				const style = derive(base);
				if (cell) {
					sheet.cells.set(key, { ...cell, style: style === 0 ? undefined : style });
				} else if (style !== 0) {
					sheet.cells.set(key, { value: null, style });
				}
			}
		}
	}
	sheet.dirty = true;
}

/**
 * Apply a style change to ranges (see restyleRanges). Identical derived
 * formats are shared.
 */
export function applyStylePatch(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
	patch: StylePatch | ((current: number | undefined) => StylePatch),
): void {
	const pool = new StylePool(workbook.styles);
	restyleRanges(tx, workbook, sheetIndex, ranges, (base) =>
		pool.derive(base, typeof patch === 'function' ? patch(base) : patch),
	);
}

function lineStyle(sheet: Worksheet, row: number, col: number): number | undefined {
	const r = sheet.rows.get(row);
	if (r?.customFormat && r.style !== undefined) {
		return r.style;
	}
	for (const c of sheet.cols) {
		if (col >= c.min && col <= c.max) {
			return c.style;
		}
	}
	return undefined;
}

function applyToWholeLines(
	sheet: Worksheet,
	range: RangeAddress,
	derive: (base: number | undefined) => number,
): void {
	// Whole columns, including select-all: column formats (one <col> record), not a million rows.
	if (isWholeColumns(range)) {
		updateCols(sheet, range.left, range.right, (info) => ({ ...info, style: derive(info.style) }));
		for (const [row, info] of sheet.rows) {
			if (info.customFormat) {
				// Row-formatted cells at the intersection need their own cell record.
				for (let col = range.left; col <= range.right; col++) {
					const key = cellKey(row, col);
					if (!sheet.cells.has(key)) {
						sheet.cells.set(key, { value: null, style: derive(info.style) });
					}
				}
			}
		}
	} else {
		for (let row = range.top; row <= range.bottom; row++) {
			updateRow(sheet, row, (info) => ({ ...info, style: derive(info.style), customFormat: true }));
		}
	}
	for (const [key, cell] of sheet.cells) {
		const row = keyRow(key);
		const col = keyCol(key);
		if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) {
			const style = derive(cell.style);
			sheet.cells.set(key, { ...cell, style: style === 0 ? undefined : style });
		}
	}
}

export type BorderPreset =
	| 'none'
	| 'all'
	| 'outside'
	| 'thickOutside'
	| 'inside'
	| 'top'
	| 'bottom'
	| 'left'
	| 'right'
	| 'bottomDouble'
	| 'thickBottom'
	| 'topBottom'
	| 'topThickBottom'
	| 'topDoubleBottom';

/**
 * Excel's border presets (Home > Borders). Edges are applied per cell so
 * that outside / inside semantics hold for multi-cell ranges.
 */
export function applyBorderPreset(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	preset: BorderPreset,
	edge: BorderEdge = { style: 'thin' },
): void {
	const sheet = workbook.sheets[sheetIndex];
	const pool = new StylePool(workbook.styles);
	tx.touchRange(sheetIndex, range, false);
	const none: BorderEdge = { style: 'none' };
	const thick: BorderEdge = { ...edge, style: 'thick' };
	const double: BorderEdge = { ...edge, style: 'double' };
	for (let row = range.top; row <= range.bottom; row++) {
		for (let col = range.left; col <= range.right; col++) {
			const top = row === range.top;
			const bottom = row === range.bottom;
			const left = col === range.left;
			const right = col === range.right;
			let b: Partial<BorderSpec> | undefined;
			switch (preset) {
				case 'none':
					b = { top: none, bottom: none, left: none, right: none, diagonal: none };
					break;
				case 'all':
					b = { top: edge, bottom: edge, left: edge, right: edge };
					break;
				case 'outside':
				case 'thickOutside': {
					const e = preset === 'thickOutside' ? thick : edge;
					b = {
						top: top ? e : undefined,
						bottom: bottom ? e : undefined,
						left: left ? e : undefined,
						right: right ? e : undefined,
					};
					break;
				}
				case 'inside':
					b = {
						top: top ? undefined : edge,
						bottom: bottom ? undefined : edge,
						left: left ? undefined : edge,
						right: right ? undefined : edge,
					};
					break;
				case 'top':
					b = top ? { top: edge } : undefined;
					break;
				case 'bottom':
					b = bottom ? { bottom: edge } : undefined;
					break;
				case 'left':
					b = left ? { left: edge } : undefined;
					break;
				case 'right':
					b = right ? { right: edge } : undefined;
					break;
				case 'bottomDouble':
					b = bottom ? { bottom: double } : undefined;
					break;
				case 'thickBottom':
					b = bottom ? { bottom: thick } : undefined;
					break;
				case 'topBottom':
					b = { top: top ? edge : undefined, bottom: bottom ? edge : undefined };
					break;
				case 'topThickBottom':
					b = { top: top ? edge : undefined, bottom: bottom ? thick : undefined };
					break;
				case 'topDoubleBottom':
					b = { top: top ? edge : undefined, bottom: bottom ? double : undefined };
					break;
			}
			if (!b || Object.values(b).every((v) => v === undefined)) {
				continue;
			}
			const key = cellKey(row, col);
			const cell = sheet.cells.get(key);
			const style = pool.derive(cell?.style ?? lineStyle(sheet, row, col), { border: b });
			sheet.cells.set(key, cell ? { ...cell, style } : { value: null, style });
		}
	}
	sheet.dirty = true;
}

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import type { RangeAddress } from '../../types/address';
import type { Cell, CellValue } from '../../types/cell';
import type { Worksheet } from '../../types/sheet';
import type { FormulaEnv } from './formula';
import type { XlsSharedStrings } from './sst';
import type { XlsWarnings } from './warnings';
import type { XlsStyles } from './xls-styles';

/** State shared by the record handlers of one worksheet substream. */
export interface SheetContext {
	sheet: Worksheet;
	styles: XlsStyles;
	sst: XlsSharedStrings;
	env: FormulaEnv;
	warnings: XlsWarnings;
}

/** Warn once per sheet and code. */
export function sheetWarn(ctx: SheetContext, code: string, message: string, cell?: string): void {
	ctx.warnings.warn(code, message, { sheet: ctx.sheet.name, cell });
}

/** A cell with a value and (unless default) a style. */
export function makeCell(value: CellValue, style: number): Cell {
	return style ? { value, style } : { value };
}

/** The cell position lies inside the SpreadsheetML grid. */
export function validCell(row: number, col: number): boolean {
	return row >= 0 && row <= MAX_ROW_INDEX && col >= 0 && col <= MAX_COL_INDEX;
}

/** A normalised range from BIFF corners, or `undefined` when outside the grid. */
export function makeRange(
	r1: number,
	r2: number,
	c1: number,
	c2: number,
): RangeAddress | undefined {
	const range = {
		top: Math.min(r1, r2),
		left: Math.min(c1, c2),
		bottom: Math.max(r1, r2),
		right: Math.max(c1, c2),
	};
	return validCell(range.top, range.left) && validCell(range.bottom, range.right)
		? range
		: undefined;
}

export function rangeHas(r: RangeAddress, row: number, col: number): boolean {
	return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

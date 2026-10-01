import { usedRange } from '../sheet/worksheet';
import { cellKey } from '../address/cell-address';
import { formatValue } from '../numfmt';
import { StyleResolver } from '../styles/resolve-style';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';

/** The text a cell displays (number formats applied). */
export function displayText(
	cell: Cell | undefined,
	styles: StyleResolver,
	date1904: boolean,
): string {
	if (!cell || cell.value === null) {
		return '';
	}
	const v = cell.value;
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (isCellError(v)) {
		return v.code;
	}
	return formatValue(v, styles.resolve(cell.style).numFmt, { date1904 }).text;
}

/** The used range of a sheet as rows of displayed text. */
export function sheetToTextGrid(
	workbook: Workbook,
	sheet: Worksheet,
	options: { range?: RangeAddress; raw?: boolean } = {},
): string[][] {
	const range = options.range ?? usedRange(sheet);
	if (!range) {
		return [];
	}
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const rows: string[][] = [];
	for (let r = range.top; r <= range.bottom; r++) {
		const row: string[] = [];
		for (let c = range.left; c <= range.right; c++) {
			const cell = sheet.cells.get(cellKey(r, c));
			if (options.raw) {
				const v = cell?.value;
				row.push(v === null || v === undefined ? '' : isCellError(v) ? v.code : String(v));
			} else {
				row.push(displayText(cell, styles, workbook.date1904));
			}
		}
		rows.push(row);
	}
	return rows;
}

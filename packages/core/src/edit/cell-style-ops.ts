import { applyNamedStyleToXf, ensureCellStyle } from '../styles/builtin-cell-styles';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { restyleRanges } from './style-ops';
import type { EditTransaction } from './transaction';

/**
 * Home > Cell Styles: apply a named style (built-in or from the workbook)
 * to ranges. The aspects the style includes replace the cells' own
 * formatting; the others are kept. Built-in styles missing from the
 * workbook are added (stylesheet entries are only appended).
 */
export function applyCellStyle(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
	styleName: string,
): void {
	const styleXfId = ensureCellStyle(workbook.styles, styleName, workbook.theme);
	if (styleXfId === undefined) {
		throw new Error(`Unknown cell style: ${styleName}`);
	}
	restyleRanges(tx, workbook, sheetIndex, ranges, (base) =>
		applyNamedStyleToXf(workbook.styles, base, styleXfId),
	);
}

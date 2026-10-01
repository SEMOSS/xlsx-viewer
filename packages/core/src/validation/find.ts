import type { DataValidation } from '../types/data-validation';
import type { Worksheet } from '../types/sheet';

/** The data validation covering a cell, if any (the first match; Excel forbids overlaps). */
export function findValidation(
	sheet: Worksheet,
	row: number,
	col: number,
): DataValidation | undefined {
	for (const dv of sheet.dataValidations ?? []) {
		for (const r of dv.ranges) {
			if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
				return dv;
			}
		}
	}
	return undefined;
}

/**
 * The input message ("prompt") to show when a validated cell is selected:
 * only when the rule enables it (`showInputMessage: true`) and has text.
 */
export function getValidationInputMessage(
	dv: DataValidation,
): { title?: string; message: string } | undefined {
	if (dv.showInputMessage !== true || (!dv.prompt && !dv.promptTitle)) {
		return undefined;
	}
	const out: { title?: string; message: string } = { message: dv.prompt ?? '' };
	if (dv.promptTitle) {
		out.title = dv.promptTitle;
	}
	return out;
}

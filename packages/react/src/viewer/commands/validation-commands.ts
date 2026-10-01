import type { DataValidationResult } from 'xlsx-viewer-core';
import { findValidation, parseCellInput, validateValue } from 'xlsx-viewer-core';

import type { EditState } from '../state/types';
import type { ViewerStore } from '../state/viewer-store';

/** Payload of the data validation alert shown for a rejected entry. */
export interface ValidationAlertPayload {
	style: NonNullable<DataValidationResult['style']>;
	title: string;
	message: string;
	/** How the edit was being committed (to finish it on "Yes" / "OK"). */
	move: 'up' | 'down' | 'left' | 'right' | 'none';
}

/**
 * Check the entry being committed against the cell's data validation.
 * Returns the failure to show, or null when the value is accepted
 * (valid, no rule, a formula, or a rule without an error alert).
 */
export function checkEntry(store: ViewerStore, edit: EditState): DataValidationResult | null {
	const sheet = store.workbook.sheets[edit.sheet];
	const dv = findValidation(sheet, edit.row, edit.col);
	if (!dv || dv.type === 'none') {
		return null;
	}
	const parsed = parseCellInput(edit.text, { date1904: store.workbook.date1904 });
	// Like Excel: formulas and clearing a cell are not validated on entry.
	if (parsed.kind === 'formula' || parsed.kind === 'blank') {
		return null;
	}
	const result = validateValue(parsed.value, dv, {
		evaluate: (f, s, r, c) => store.calc.evaluateFormula(f, s, r, c),
		sheetIndex: edit.sheet,
		row: edit.row,
		col: edit.col,
		date1904: store.workbook.date1904,
		workbook: store.workbook,
	});
	return result.valid || !result.style ? null : result;
}

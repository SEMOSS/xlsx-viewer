import type { SheetProtection } from '../types/sheet';
import type { Workbook, WorkbookProtection } from '../types/workbook';
import type { EditTransaction } from './transaction';

/** Review > Protect Sheet: replace the sheet's protection (hashes are computed by the caller). */
export function protectSheet(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	protection: SheetProtection,
): void {
	tx.touchSheet(sheetIndex, false, 'view');
	const sheet = workbook.sheets[sheetIndex];
	sheet.protection = { ...protection, sheet: true };
	sheet.dirty = true;
}

/** Review > Unprotect Sheet (the caller checks the password). */
export function unprotectSheet(tx: EditTransaction, workbook: Workbook, sheetIndex: number): void {
	tx.touchSheet(sheetIndex, false, 'view');
	const sheet = workbook.sheets[sheetIndex];
	sheet.protection = undefined;
	sheet.dirty = true;
}

/** Review > Protect Workbook (structure), or remove it with `undefined`. */
export function setWorkbookProtection(
	tx: EditTransaction,
	workbook: Workbook,
	protection: WorkbookProtection | undefined,
): void {
	tx.touchWorkbook('names');
	workbook.protection = protection;
	workbook.dirty = true;
}

import { cellKey, keyCol, keyRow } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import type { SheetProtection, Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { CellChange, SheetChange, TransactionGuard, WorkbookChange } from './transaction';

/** Actions a protected sheet can allow (OOXML `sheetProtection` attributes). */
export type ProtectableAction = Exclude<
	keyof SheetProtection,
	'password' | 'algorithmName' | 'hashValue' | 'saltValue' | 'spinCount' | 'sheet'
>;

/** Schema defaults: `true` means the action is locked while the sheet is protected. */
const LOCKED_BY_DEFAULT: Record<ProtectableAction, boolean> = {
	objects: false,
	scenarios: false,
	formatCells: true,
	formatColumns: true,
	formatRows: true,
	insertColumns: true,
	insertRows: true,
	insertHyperlinks: true,
	deleteColumns: true,
	deleteRows: true,
	selectLockedCells: false,
	sort: true,
	autoFilter: true,
	pivotTables: true,
	selectUnlockedCells: false,
};

export const PROTECTED_SHEET_MESSAGE =
	"The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.";
export const PROTECTED_WORKBOOK_MESSAGE = 'Workbook is protected and cannot be changed.';

/** A change blocked by sheet or workbook protection. */
export class ProtectedError extends Error {
	constructor(
		message: string,
		readonly reason: 'locked-cells' | 'sheet' | 'workbook',
	) {
		super(message);
		this.name = 'ProtectedError';
	}
}

export function isSheetProtected(sheet: Worksheet): boolean {
	return sheet.protection?.sheet === true;
}

/** The action is possible on this sheet (always true when the sheet is not protected). */
export function isActionAllowed(sheet: Worksheet, action: ProtectableAction): boolean {
	if (!isSheetProtected(sheet)) {
		return true;
	}
	return !(sheet.protection?.[action] ?? LOCKED_BY_DEFAULT[action]);
}

function styleIdAt(sheet: Worksheet, row: number, col: number): number {
	const cell = sheet.cells.get(cellKey(row, col));
	if (cell?.style !== undefined) {
		return cell.style;
	}
	const info = sheet.rows.get(row);
	if (info?.customFormat && info.style !== undefined) {
		return info.style;
	}
	return sheet.cols.find((c) => col >= c.min && col <= c.max)?.style ?? 0;
}

function lockedXf(workbook: Workbook, xfId: number): boolean {
	return workbook.styles.cellXfs[xfId]?.protection?.locked !== false;
}

/** The cell is locked (the default for every cell; Format Cells > Protection). */
export function isCellLocked(
	workbook: Workbook,
	sheet: Worksheet,
	row: number,
	col: number,
): boolean {
	return lockedXf(workbook, styleIdAt(sheet, row, col));
}

/**
 * Some cell of the range is locked. Large ranges check their existing
 * cells plus the row, column and default formats that the empty cells use.
 */
export function rangeHasLockedCells(
	workbook: Workbook,
	sheet: Worksheet,
	range: RangeAddress,
): boolean {
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	if (area <= 50_000) {
		for (let row = range.top; row <= range.bottom; row++) {
			for (let col = range.left; col <= range.right; col++) {
				if (isCellLocked(workbook, sheet, row, col)) {
					return true;
				}
			}
		}
		return false;
	}
	for (const [key, cell] of sheet.cells) {
		const row = keyRow(key);
		const col = keyCol(key);
		if (
			row >= range.top &&
			row <= range.bottom &&
			col >= range.left &&
			col <= range.right &&
			cell.style !== undefined &&
			lockedXf(workbook, cell.style)
		) {
			return true;
		}
	}
	for (const [row, info] of sheet.rows) {
		if (
			row >= range.top &&
			row <= range.bottom &&
			info.customFormat &&
			lockedXf(workbook, info.style ?? 0)
		) {
			return true;
		}
	}
	for (const c of sheet.cols) {
		if (c.max >= range.left && c.min <= range.right && lockedXf(workbook, c.style ?? 0)) {
			return true;
		}
	}
	// Columns of the range without a column format use the default format.
	let covered = 0;
	for (const c of sheet.cols) {
		if (c.style !== undefined) {
			covered += Math.max(0, Math.min(c.max, range.right) - Math.max(c.min, range.left) + 1);
		}
	}
	return covered < range.right - range.left + 1 && lockedXf(workbook, 0);
}

const SHEET_CHANGE_ACTION: Partial<Record<SheetChange, ProtectableAction>> = {
	insertRows: 'insertRows',
	insertColumns: 'insertColumns',
	deleteRows: 'deleteRows',
	deleteColumns: 'deleteColumns',
	formatColumns: 'formatColumns',
	formatRows: 'formatRows',
	filter: 'autoFilter',
	sort: 'sort',
	hyperlinks: 'insertHyperlinks',
	objects: 'objects',
};

/**
 * Transaction guard enforcing sheet and workbook protection like Excel:
 * values only change in unlocked cells, other changes need the matching
 * "Allow all users of this worksheet to" option, and a workbook with a
 * protected structure keeps its sheet list.
 */
export function createProtectionGuard(workbook: Workbook): TransactionGuard {
	const blocked = () => new ProtectedError(PROTECTED_SHEET_MESSAGE, 'sheet');
	return {
		cells(sheetIndex: number, range: RangeAddress, change: CellChange): void {
			const sheet = workbook.sheets[sheetIndex];
			if (!sheet || !isSheetProtected(sheet)) {
				return;
			}
			if (change === 'format') {
				if (!isActionAllowed(sheet, 'formatCells')) {
					throw blocked();
				}
			} else if (rangeHasLockedCells(workbook, sheet, range)) {
				throw new ProtectedError(PROTECTED_SHEET_MESSAGE, 'locked-cells');
			}
		},
		sheet(sheetIndex: number, change: SheetChange): void {
			const sheet = workbook.sheets[sheetIndex];
			if (
				!sheet ||
				!isSheetProtected(sheet) ||
				change === 'view' ||
				change === 'references' ||
				change === 'sheets'
			) {
				return;
			}
			const action = SHEET_CHANGE_ACTION[change];
			if (!action || !isActionAllowed(sheet, action)) {
				throw blocked();
			}
		},
		workbook(change: WorkbookChange): void {
			if (change === 'sheets' && workbook.protection?.lockStructure) {
				throw new ProtectedError(PROTECTED_WORKBOOK_MESSAGE, 'workbook');
			}
		},
	};
}

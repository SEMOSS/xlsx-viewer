import type { ProtectableAction, ProtectionHash, WorkbookProtection } from 'xlsx-viewer-core';
import {
	createProtectionHash,
	hasProtectionPassword,
	isCellLocked,
	isSheetProtected,
	PROTECTED_SHEET_MESSAGE,
	protectSheet,
	setWorkbookProtection,
	unprotectSheet,
	verifyProtectionPassword,
} from 'xlsx-viewer-core';

import type { ShellActions } from '../shell-context';
import type { ViewerStore } from '../state/viewer-store';

/** Show Excel's alert for a change blocked by protection. */
export function alertProtected(store: ViewerStore, message = PROTECTED_SHEET_MESSAGE): void {
	store.setUi({ dialog: { kind: 'alert', payload: { message } } });
}

/** The active cell can't be edited: locked on a protected sheet (shows the alert). */
export function blockedByProtection(store: ViewerStore, row: number, col: number): boolean {
	const sheet = store.activeSheet;
	if (isSheetProtected(sheet) && isCellLocked(store.workbook, sheet, row, col)) {
		alertProtected(store);
		return true;
	}
	return false;
}

/** Review > Protect Sheet. `locked` maps actions to OOXML flags (true = not allowed). */
export async function protectSheetCmd(
	store: ViewerStore,
	password: string,
	locked: Partial<Record<ProtectableAction, boolean>>,
): Promise<void> {
	const hash = password ? await createProtectionHash(password) : {};
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.protectSheet'), (tx) =>
		protectSheet(tx, store.workbook, sheetIndex, { sheet: true, ...locked, ...hash }),
	);
}

function workbookHash(p: WorkbookProtection | undefined): ProtectionHash {
	return {
		password: p?.workbookPassword,
		algorithmName: p?.workbookAlgorithmName,
		hashValue: p?.workbookHashValue,
		saltValue: p?.workbookSaltValue,
		spinCount: p?.workbookSpinCount,
	};
}

/** Review > Protect Workbook (structure). */
export async function protectWorkbookCmd(store: ViewerStore, password: string): Promise<void> {
	const hash = password ? await createProtectionHash(password) : undefined;
	store.transact(store.t('protect.protectWorkbook'), (tx) =>
		setWorkbookProtection(tx, store.workbook, {
			lockStructure: true,
			...(hash && {
				workbookAlgorithmName: hash.algorithmName,
				workbookHashValue: hash.hashValue,
				workbookSaltValue: hash.saltValue,
				workbookSpinCount: hash.spinCount,
			}),
		}),
	);
}

/** Remove sheet or workbook protection after checking the password. Returns false when it is wrong. */
export async function unprotectCmd(
	store: ViewerStore,
	target: 'sheet' | 'workbook',
	password: string,
): Promise<boolean> {
	const sheetIndex = store.state.activeSheet;
	if (target === 'sheet') {
		const p = store.activeSheet.protection;
		if (p && !(await verifyProtectionPassword(p, password))) {
			return false;
		}
		store.transact(store.t('protect.unprotectSheet'), (tx) =>
			unprotectSheet(tx, store.workbook, sheetIndex),
		);
		return true;
	}
	if (!(await verifyProtectionPassword(workbookHash(store.workbook.protection), password))) {
		return false;
	}
	store.transact(store.t('protect.unprotectWorkbook'), (tx) =>
		setWorkbookProtection(tx, store.workbook, undefined),
	);
	return true;
}

/** Review > Protect Sheet / Unprotect Sheet button. */
export function toggleSheetProtection(store: ViewerStore, shell: ShellActions): void {
	const p = store.activeSheet.protection;
	if (!p?.sheet) {
		shell.openDialog('protectSheet');
	} else if (hasProtectionPassword(p)) {
		shell.openDialog('unprotect', { target: 'sheet' });
	} else {
		void unprotectCmd(store, 'sheet', '');
	}
}

/** Review > Protect Workbook button (toggles structure protection). */
export function toggleWorkbookProtection(store: ViewerStore, shell: ShellActions): void {
	const p = store.workbook.protection;
	if (!p?.lockStructure) {
		shell.openDialog('protectWorkbook');
	} else if (hasProtectionPassword(workbookHash(p))) {
		shell.openDialog('unprotect', { target: 'workbook' });
	} else {
		void unprotectCmd(store, 'workbook', '');
	}
}

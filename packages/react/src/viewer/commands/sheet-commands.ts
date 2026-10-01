import type { ColorSpec } from 'xlsx-viewer-core';
import {
	addSheet,
	deleteSheet,
	duplicateSheet,
	moveSheet,
	renameSheet,
	setSheetTabColor,
	setSheetVisibility,
	SheetNameError,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

/** Insert a new sheet after the active one and activate it. */
export function insertSheetCmd(store: ViewerStore, index?: number): void {
	let at = -1;
	store.transact(store.t('cmd.insertSheet'), (tx) => {
		at = addSheet(tx, store.workbook, index);
	});
	if (at >= 0) {
		store.syncSheets();
		store.setUi({ activeSheet: at });
	}
}

export function deleteSheetCmd(store: ViewerStore, index = store.state.activeSheet): void {
	try {
		store.transact(store.t('cmd.deleteSheet'), (tx) => deleteSheet(tx, store.workbook, index));
		store.syncSheets();
		const active = Math.min(index, store.workbook.sheets.length - 1);
		store.setUi({ activeSheet: active });
	} catch (e) {
		store.setUi({ message: (e as Error).message });
	}
}

/** Returns an error message when the name is rejected. */
export function renameSheetCmd(store: ViewerStore, index: number, name: string): string | null {
	try {
		store.transact(store.t('cmd.renameSheet'), (tx) =>
			renameSheet(tx, store.workbook, index, name.trim()),
		);
		return null;
	} catch (e) {
		if (e instanceof SheetNameError) {
			return store.t(
				e.reason === 'duplicate' ? 'message.sheetNameTaken' : 'message.sheetNameInvalid',
			);
		}
		throw e;
	}
}

export function moveSheetCmd(store: ViewerStore, from: number, to: number): void {
	if (from === to || to < 0 || to >= store.workbook.sheets.length) {
		return;
	}
	store.transact(store.t('undo.moveSheet'), (tx) => moveSheet(tx, store.workbook, from, to));
	store.setUi({ activeSheet: to });
}

export function duplicateSheetCmd(store: ViewerStore, index: number): void {
	let at = -1;
	store.transact(store.t('undo.copySheet'), (tx) => {
		at = duplicateSheet(tx, store.workbook, index);
	});
	if (at >= 0) {
		store.syncSheets();
		store.setUi({ activeSheet: at });
	}
}

export function setTabColorCmd(
	store: ViewerStore,
	index: number,
	color: ColorSpec | undefined,
): void {
	store.transact(store.t('cmd.tabColor'), (tx) =>
		setSheetTabColor(tx, store.workbook, index, color),
	);
}

export function hideSheetCmd(store: ViewerStore, index: number, hidden: boolean): void {
	try {
		store.transact(store.t(hidden ? 'cmd.hideSheet' : 'undo.unhideSheet'), (tx) =>
			setSheetVisibility(tx, store.workbook, index, hidden ? 'hidden' : 'visible'),
		);
		if (hidden && store.state.activeSheet === index) {
			const next = store.workbook.sheets.findIndex((s) => s.state === 'visible');
			store.setUi({ activeSheet: Math.max(0, next) });
		} else if (!hidden) {
			store.setUi({ activeSheet: index });
		}
	} catch (e) {
		store.setUi({ message: (e as Error).message });
	}
}

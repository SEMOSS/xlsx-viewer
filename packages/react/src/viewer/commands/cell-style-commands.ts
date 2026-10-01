import { applyCellStyle } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

/** Home > Cell Styles: apply a named style to the selection. */
export function applyCellStyleCmd(store: ViewerStore, styleName: string): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('undo.cellStyle', { name: styleName }), (tx) => {
		applyCellStyle(tx, store.workbook, sheetIndex, store.selectionOf().ranges, styleName);
	});
}

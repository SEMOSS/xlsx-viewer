import { setFreezePanes, setSheetZoom } from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

export type FreezeMode = 'panes' | 'topRow' | 'firstColumn' | 'none';

/** View > Freeze Panes. */
export function freezePanes(
	store: ViewerStore,
	mode: FreezeMode,
	topLeftVisible?: { row: number; col: number },
): void {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const frozen = sheet.view.pane?.state === 'frozen';
	const active = store.selectionOf().active;
	store.transact(store.t('cmd.freezePanes'), (tx) => {
		if (mode === 'none' || (mode === 'panes' && frozen)) {
			setFreezePanes(tx, store.workbook, sheetIndex, null);
		} else if (mode === 'topRow') {
			setFreezePanes(
				tx,
				store.workbook,
				sheetIndex,
				{ row: (topLeftVisible?.row ?? 0) + 1, col: 0 },
				{ row: topLeftVisible?.row ?? 0, col: 0 },
			);
		} else if (mode === 'firstColumn') {
			setFreezePanes(
				tx,
				store.workbook,
				sheetIndex,
				{ row: 0, col: (topLeftVisible?.col ?? 0) + 1 },
				{ row: 0, col: topLeftVisible?.col ?? 0 },
			);
		} else {
			setFreezePanes(tx, store.workbook, sheetIndex, active, topLeftVisible);
		}
	});
}

export function setZoom(store: ViewerStore, percent: number): void {
	setSheetZoom(store.workbook, store.state.activeSheet, percent);
	store.touch();
}

/** Zoom so that the selection fills the visible grid area. */
export function zoomToSelection(
	store: ViewerStore,
	viewport: { width: number; height: number },
	rangePx: { width: number; height: number },
): void {
	const zoom = (store.activeSheet.view.zoomScale || 100) / 100;
	const w = rangePx.width / zoom;
	const h = rangePx.height / zoom;
	const fit = Math.min(viewport.width / Math.max(1, w), viewport.height / Math.max(1, h));
	setZoom(store, Math.max(10, Math.min(400, Math.floor(fit * 100))));
	void currentRange;
}

export function toggleSheetViewFlag(
	store: ViewerStore,
	flag: 'showGridLines' | 'showRowColHeaders' | 'showFormulas' | 'showZeros',
): void {
	const sheet = store.activeSheet;
	sheet.view = { ...sheet.view, [flag]: !sheet.view[flag] };
	sheet.dirty = true;
	store.touch();
}

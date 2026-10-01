import type {
	EditTransaction,
	MarginPreset,
	PageMargins,
	Workbook,
	Worksheet,
} from 'xlsx-viewer-core';
import {
	addToPrintArea,
	canRemovePageBreak,
	clearPrintArea,
	insertPageBreak,
	isWholeColumnRange,
	isWholeRowRange,
	marginPresetOf,
	removePageBreak,
	resetAllPageBreaks,
	setFitTo,
	setMargins,
	setOrientation,
	setPaperSize,
	setPrintArea,
	setPrintGridlines,
	setPrintHeadings,
	setScale,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../../state/viewer-store';

/**
 * Page Layout commands on the active sheet, each one undoable step (the
 * ribbon tab, File > Print and the Page Setup dialog share them).
 */

export type ScalingMode = 'none' | 'fitSheet' | 'fitColumns' | 'fitRows' | 'custom';

/** Run a page setup change on the active sheet as one undoable step (`label`: translated). */
export function pageSetupStep(
	store: ViewerStore,
	label: string,
	fn: (tx: EditTransaction, workbook: Workbook, sheetIndex: number) => void,
): boolean {
	const sheetIndex = store.state.activeSheet;
	try {
		return store.transact(label, (tx) => fn(tx, store.workbook, sheetIndex), {
			autoRowHeight: false,
		});
	} catch (e) {
		store.setUi({ message: (e as Error).message });
		return false;
	}
}

export function orientationCmd(store: ViewerStore, orientation: 'portrait' | 'landscape'): void {
	pageSetupStep(store, store.t('cmd.orientation'), (tx, wb, i) =>
		setOrientation(tx, wb, i, orientation),
	);
}

export function paperSizeCmd(store: ViewerStore, paperSize: number): void {
	pageSetupStep(store, store.t('cmd.size'), (tx, wb, i) => setPaperSize(tx, wb, i, paperSize));
}

export function marginsCmd(store: ViewerStore, margins: MarginPreset | PageMargins): void {
	pageSetupStep(store, store.t('cmd.margins'), (tx, wb, i) => setMargins(tx, wb, i, margins));
}

/** The preset the sheet's margins match ('custom' otherwise). */
export function marginsMode(sheet: Worksheet): MarginPreset | 'custom' {
	return marginPresetOf(sheet.pageMargins) ?? 'custom';
}

/** The scaling choice of File > Print for a sheet. */
export function scalingMode(sheet: Worksheet): ScalingMode {
	const ps = sheet.pageSetup;
	if (ps?.fitToPage) {
		const w = ps.fitToWidth ?? 1;
		const h = ps.fitToHeight ?? 1;
		if (w === 1 && h === 1) {
			return 'fitSheet';
		}
		if (w === 1 && h === 0) {
			return 'fitColumns';
		}
		if (w === 0 && h === 1) {
			return 'fitRows';
		}
		return w === 0 && h === 0 ? 'none' : 'custom';
	}
	return (ps?.scale ?? 100) === 100 ? 'none' : 'custom';
}

export function scalingCmd(store: ViewerStore, mode: Exclude<ScalingMode, 'custom'>): void {
	pageSetupStep(store, store.t('print.scaling'), (tx, wb, i) => {
		if (mode === 'none') {
			setScale(tx, wb, i, 100);
		} else if (mode === 'fitSheet') {
			setFitTo(tx, wb, i, 1, 1);
		} else if (mode === 'fitColumns') {
			setFitTo(tx, wb, i, 1, 0);
		} else {
			setFitTo(tx, wb, i, 0, 1);
		}
	});
}

/** Scale to Fit > Width / Height (0 = Automatic); keeps the other dimension. */
export function fitToCmd(store: ViewerStore, axis: 'width' | 'height', pages: number): void {
	const ps = store.activeSheet.pageSetup;
	const fitting = ps?.fitToPage === true;
	const width = axis === 'width' ? pages : fitting ? (ps?.fitToWidth ?? 1) : 0;
	const height = axis === 'height' ? pages : fitting ? (ps?.fitToHeight ?? 1) : 0;
	pageSetupStep(store, store.t('group.scaleToFit'), (tx, wb, i) =>
		setFitTo(tx, wb, i, width, height),
	);
}

export function scaleCmd(store: ViewerStore, percent: number): void {
	pageSetupStep(store, store.t('cmd.scale'), (tx, wb, i) => setScale(tx, wb, i, percent));
}

export function printAreaCmd(store: ViewerStore, action: 'set' | 'add' | 'clear'): void {
	const ranges = store.selectionOf().ranges;
	pageSetupStep(store, store.t('cmd.printArea'), (tx, wb, i) => {
		if (action === 'clear') {
			clearPrintArea(tx, wb, i);
		} else if (action === 'add') {
			addToPrintArea(tx, wb, i, ranges);
		} else {
			setPrintArea(tx, wb, i, ranges);
		}
	});
}

/** Insert / Remove Page Break at the active cell; whole rows (columns) only break that axis. */
export function pageBreakCmd(store: ViewerStore, action: 'insert' | 'remove' | 'reset'): void {
	const sel = store.selectionOf();
	const range = sel.ranges[0];
	const axes =
		range && isWholeRowRange(range) ? 'row' : range && isWholeColumnRange(range) ? 'col' : 'both';
	const label = store.t(
		action === 'insert'
			? 'pageLayout.insertBreak'
			: action === 'remove'
				? 'pageLayout.removeBreak'
				: 'pageLayout.resetBreaks',
	);
	pageSetupStep(store, label, (tx, wb, i) => {
		if (action === 'insert') {
			insertPageBreak(tx, wb, i, sel.active, axes);
		} else if (action === 'remove') {
			removePageBreak(tx, wb, i, sel.active, axes);
		} else {
			resetAllPageBreaks(tx, wb, i);
		}
	});
}

/** Remove Page Break is possible at the active cell. */
export function canRemoveBreakAtActive(store: ViewerStore): boolean {
	return canRemovePageBreak(store.activeSheet, store.selectionOf().active);
}

export function printFlagCmd(
	store: ViewerStore,
	flag: 'gridLines' | 'headings',
	on: boolean,
): void {
	const label = store.t(flag === 'gridLines' ? 'undo.printGridlines' : 'undo.printHeadings');
	pageSetupStep(store, label, (tx, wb, i) =>
		flag === 'gridLines' ? setPrintGridlines(tx, wb, i, on) : setPrintHeadings(tx, wb, i, on),
	);
}

import { cellKey, formatRange, localDateToSerial } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';
import { beginEdit, updateEditText } from './edit-commands';

function isNumberCell(store: ViewerStore, row: number, col: number): boolean {
	const c = store.activeSheet.cells.get(cellKey(row, col));
	return typeof c?.value === 'number';
}

/**
 * AutoSum (Alt+=): propose `=FUNC(range)` over the numbers directly above
 * the active cell (or to its left) and start editing it.
 */
export function autoSum(store: ViewerStore, fn = 'SUM'): void {
	if (store.readOnly) {
		return;
	}
	const { row, col } = store.selectionOf().active;
	let range: string | null = null;
	if (row > 0 && isNumberCell(store, row - 1, col)) {
		let top = row - 1;
		while (top > 0 && isNumberCell(store, top - 1, col)) {
			top--;
		}
		range = formatRange({ top, left: col, bottom: row - 1, right: col });
	} else if (col > 0 && isNumberCell(store, row, col - 1)) {
		let left = col - 1;
		while (left > 0 && isNumberCell(store, row, left - 1)) {
			left--;
		}
		range = formatRange({ top: row, left, bottom: row, right: col - 1 });
	}
	const text = `=${fn}(${range ?? ''})`;
	beginEdit(store, { initialText: text, mode: 'enter' });
	const edit = store.state.edit;
	if (edit && !range) {
		store.setUi({ edit: { ...edit, caret: text.length - 1 } });
	}
}

/** Ctrl+; / Ctrl+Shift+; : start typing today's date / the current time. */
export function insertNow(store: ViewerStore, part: 'date' | 'time'): void {
	const now = new Date();
	const text =
		part === 'date'
			? `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`
			: `${now.getHours() % 12 || 12}:${String(now.getMinutes()).padStart(2, '0')} ${now.getHours() < 12 ? 'AM' : 'PM'}`;
	void localDateToSerial;
	const edit = store.state.edit;
	if (edit) {
		const next = edit.text.slice(0, edit.caret) + text + edit.text.slice(edit.caret);
		store.setUi({ edit: { ...edit, text: next, caret: edit.caret + text.length } });
	} else {
		beginEdit(store, { initialText: text, mode: 'enter' });
	}
}

/** Ctrl+PageUp / Ctrl+PageDown: switch to the previous / next visible sheet. */
export function switchSheet(store: ViewerStore, delta: 1 | -1): void {
	const sheets = store.workbook.sheets;
	let i = store.state.activeSheet;
	for (let n = 0; n < sheets.length; n++) {
		i = (i + delta + sheets.length) % sheets.length;
		if (sheets[i].state === 'visible') {
			break;
		}
	}
	if (i !== store.state.activeSheet) {
		store.setUi({ activeSheet: i, edit: null });
	}
}

export function toggleShowFormulas(store: ViewerStore): void {
	const sheet = store.activeSheet;
	sheet.view = { ...sheet.view, showFormulas: !sheet.view.showFormulas };
	store.touch();
}

/** Formulas > Use in Formula: put a defined name into the formula being typed (or start one). */
export function insertNameInFormula(store: ViewerStore, name: string): void {
	if (store.readOnly) {
		return;
	}
	const edit = store.state.edit;
	if (edit && edit.text.startsWith('=')) {
		updateEditText(
			store,
			edit.text.slice(0, edit.caret) + name + edit.text.slice(edit.caret),
			edit.caret + name.length,
		);
		return;
	}
	beginEdit(store, { initialText: `=${name}`, mode: 'enter' });
}

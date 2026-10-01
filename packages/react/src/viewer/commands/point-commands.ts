import type { RangeAddress } from 'xlsx-viewer-core';
import { formatRange, quoteSheetName } from 'xlsx-viewer-core';

import type { Direction } from '../../engine/selection/navigation';
import { rangeFromCorners } from '../../engine/selection/selection';
import type { EditState } from '../state/types';
import type { ViewerStore } from '../state/viewer-store';

/** Characters after which a reference may be inserted while typing a formula. */
const POINTABLE_AFTER = new Set([
	'=',
	'(',
	',',
	';',
	'+',
	'-',
	'*',
	'/',
	'^',
	'&',
	'<',
	'>',
	':',
	'{',
	' ',
	'%',
]);

/**
 * Excel's "Point" mode: the caret is at a place in a formula where a cell
 * reference can be inserted by clicking or with the arrow keys.
 */
export function canPoint(edit: EditState): boolean {
	if (!edit.text.startsWith('=')) {
		return false;
	}
	if (edit.point && edit.caret === edit.point.end) {
		return true;
	}
	if (edit.caret !== edit.text.length && /[A-Za-z0-9_.$!'"\]]/.test(edit.text[edit.caret] ?? '')) {
		return false;
	}
	let i = edit.caret - 1;
	while (i > 0 && edit.text[i] === ' ') {
		i--;
	}
	const prev = edit.text[i];
	// Inside a string literal there is no pointing.
	const quotes = (edit.text.slice(0, edit.caret).match(/"/g) ?? []).length;
	return quotes % 2 === 0 && POINTABLE_AFTER.has(prev);
}

function refText(store: ViewerStore, edit: EditState, sheet: number, range: RangeAddress): string {
	const local = formatRange(range);
	if (sheet === edit.sheet) {
		return local;
	}
	return `${quoteSheetName(store.workbook.sheets[sheet].name)}!${local}`;
}

function applyPoint(
	store: ViewerStore,
	edit: EditState,
	sheet: number,
	anchor: { row: number; col: number },
	range: RangeAddress,
): void {
	const text = refText(store, edit, sheet, range);
	const start = edit.point && edit.caret === edit.point.end ? edit.point.start : edit.caret;
	const end = edit.point && edit.caret === edit.point.end ? edit.point.end : edit.caret;
	const nextText = edit.text.slice(0, start) + text + edit.text.slice(end);
	const caret = start + text.length;
	store.setUi({
		edit: { ...edit, text: nextText, caret, point: { range, start, end: caret, anchor, sheet } },
	});
}

/** Click on a cell while editing: insert a reference when pointing is allowed. Returns true if handled. */
export function tryPointReference(
	store: ViewerStore,
	row: number,
	col: number,
	extend: boolean,
): boolean {
	const edit = store.state.edit;
	if (!edit || !canPoint(edit)) {
		return false;
	}
	const sheet = store.state.activeSheet;
	if (extend && edit.point) {
		applyPoint(
			store,
			edit,
			sheet,
			edit.point.anchor,
			rangeFromCorners(edit.point.anchor, { row, col }),
		);
	} else {
		applyPoint(store, edit, sheet, { row, col }, { top: row, left: col, bottom: row, right: col });
	}
	return true;
}

/** Drag while pointing: extend the pointed range. */
export function pointDragTo(store: ViewerStore, row: number, col: number): void {
	const edit = store.state.edit;
	if (!edit?.point) {
		return;
	}
	applyPoint(
		store,
		edit,
		edit.point.sheet,
		edit.point.anchor,
		rangeFromCorners(edit.point.anchor, { row, col }),
	);
}

/** Arrow keys while pointing: move (or extend with Shift) the pointed reference. Returns true if handled. */
export function pointMove(store: ViewerStore, dir: Direction, extend: boolean): boolean {
	const edit = store.state.edit;
	if (!edit || edit.mode !== 'enter' || !canPoint(edit)) {
		return false;
	}
	const base = edit.point ?? {
		anchor: { row: edit.row, col: edit.col },
		range: { top: edit.row, left: edit.col, bottom: edit.row, right: edit.col },
		start: edit.caret,
		end: edit.caret,
		sheet: store.state.activeSheet,
	};
	const dr = dir === 'down' ? 1 : dir === 'up' ? -1 : 0;
	const dc = dir === 'right' ? 1 : dir === 'left' ? -1 : 0;
	if (extend && edit.point) {
		const r = base.range;
		const corner = {
			row: base.anchor.row === r.top ? r.bottom : r.top,
			col: base.anchor.col === r.left ? r.right : r.left,
		};
		const next = { row: Math.max(0, corner.row + dr), col: Math.max(0, corner.col + dc) };
		applyPoint(store, edit, base.sheet, base.anchor, rangeFromCorners(base.anchor, next));
		return true;
	}
	const startCell = edit.point ? base.anchor : { row: edit.row, col: edit.col };
	const next = { row: Math.max(0, startCell.row + dr), col: Math.max(0, startCell.col + dc) };
	applyPoint(store, edit, base.sheet, next, {
		top: next.row,
		left: next.col,
		bottom: next.row,
		right: next.col,
	});
	return true;
}

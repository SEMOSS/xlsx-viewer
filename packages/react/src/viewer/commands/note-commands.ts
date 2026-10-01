import {
	adjacentNote,
	deleteNotes,
	noteAt,
	setAllNotesVisible,
	setNoteVisible,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';
import { selectCell } from './selection-commands';

/** The note of the active cell, if any. */
export function activeNote(store: ViewerStore) {
	const { row, col } = store.selectionOf().active;
	return noteAt(store.workbook, store.state.activeSheet, row, col);
}

/** Review > Delete (notes in the selection). */
export function deleteNotesCmd(store: ViewerStore): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.deleteNote'), (tx) =>
		deleteNotes(tx, store.workbook, sheetIndex, store.selectionOf().ranges),
	);
}

/** Review > Previous / Next note: select the cell of the adjacent note. */
export function goToAdjacentNote(store: ViewerStore, direction: 1 | -1): void {
	const note = adjacentNote(
		store.workbook,
		store.state.activeSheet,
		store.selectionOf().active,
		direction,
	);
	if (note) {
		selectCell(store, note.ref.row, note.ref.col);
	}
}

/** Review > Show/Hide Note for the active cell. */
export function toggleActiveNoteVisible(store: ViewerStore): void {
	const note = activeNote(store);
	if (!note) {
		return;
	}
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t(note.visible ? 'undo.hideNote' : 'undo.showNote'), (tx) =>
		setNoteVisible(tx, store.workbook, sheetIndex, note.ref.row, note.ref.col, !note.visible),
	);
}

/** All notes of the sheet are shown. */
export function allNotesVisible(store: ViewerStore): boolean {
	const notes = store.activeSheet.notes;
	return notes.length > 0 && notes.every((n) => n.visible);
}

/** Review > Show All Notes (toggle). */
export function toggleAllNotes(store: ViewerStore): void {
	const show = !allNotesVisible(store);
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t(show ? 'cmd.showAllNotes' : 'undo.hideAllNotes'), (tx) =>
		setAllNotesVisible(tx, store.workbook, sheetIndex, show),
	);
}

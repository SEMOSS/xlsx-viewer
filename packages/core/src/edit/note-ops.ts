import type { CellAddress, RangeAddress } from '../types/address';
import type { Note } from '../types/comment';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

function begin(tx: EditTransaction, workbook: Workbook, sheetIndex: number) {
	// Notes are drawing objects for sheet protection ("Edit objects").
	tx.touchSheet(sheetIndex, false, 'objects');
	const sheet = workbook.sheets[sheetIndex];
	sheet.commentsDirty = true;
	sheet.dirty = true;
	return sheet;
}

const at = (n: Note, row: number, col: number) => n.ref.row === row && n.ref.col === col;

/** The note of a cell, if any. */
export function noteAt(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): Note | undefined {
	return workbook.sheets[sheetIndex].notes.find((n) => at(n, row, col));
}

/** New Note / Edit Note: set a cell's note text; empty text removes the note. */
export function setNote(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	text: string,
	author: string,
): void {
	const sheet = begin(tx, workbook, sheetIndex);
	const existing = sheet.notes.find((n) => at(n, row, col));
	const others = sheet.notes.filter((n) => !at(n, row, col));
	sheet.notes = text.trim()
		? [
				...others,
				{
					...(existing ?? {}),
					ref: { row, col },
					author: existing?.author ?? author,
					text,
					runs: undefined,
				},
			]
		: others;
}

/** Review > Delete: remove the notes inside ranges. Returns how many were removed. */
export function deleteNotes(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
): number {
	const inside = (n: Note) =>
		ranges.some(
			(r) =>
				n.ref.row >= r.top && n.ref.row <= r.bottom && n.ref.col >= r.left && n.ref.col <= r.right,
		);
	const count = workbook.sheets[sheetIndex].notes.filter(inside).length;
	if (count > 0) {
		const sheet = begin(tx, workbook, sheetIndex);
		sheet.notes = sheet.notes.filter((n) => !inside(n));
	}
	return count;
}

/** Show/Hide Note for one cell. */
export function setNoteVisible(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	visible: boolean,
): void {
	const sheet = begin(tx, workbook, sheetIndex);
	sheet.notes = sheet.notes.map((n) =>
		at(n, row, col) ? { ...n, visible: visible || undefined } : n,
	);
}

/** Show All Notes / hide them all again. */
export function setAllNotesVisible(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	visible: boolean,
): void {
	const sheet = begin(tx, workbook, sheetIndex);
	sheet.notes = sheet.notes.map((n) => ({ ...n, visible: visible || undefined }));
}

/** Previous / Next note in reading order (row by row), wrapping around like Excel. */
export function adjacentNote(
	workbook: Workbook,
	sheetIndex: number,
	from: CellAddress,
	direction: 1 | -1,
): Note | undefined {
	const notes = [...workbook.sheets[sheetIndex].notes].sort(
		(a, b) => a.ref.row - b.ref.row || a.ref.col - b.ref.col,
	);
	if (notes.length === 0) {
		return undefined;
	}
	const cmp = (n: Note) => n.ref.row - from.row || n.ref.col - from.col;
	if (direction === 1) {
		return notes.find((n) => cmp(n) > 0) ?? notes[0];
	}
	return [...notes].reverse().find((n) => cmp(n) < 0) ?? notes[notes.length - 1];
}

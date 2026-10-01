import { describe, expect, it } from 'vitest';

import { createWorkbook } from '../sheet/workbook';
import {
	adjacentNote,
	deleteNotes,
	noteAt,
	setAllNotesVisible,
	setNote,
	setNoteVisible,
} from './note-ops';
import { EditTransaction } from './transaction';

const run = (wb: ReturnType<typeof createWorkbook>, fn: (tx: EditTransaction) => void) => {
	const tx = new EditTransaction(wb, 'Notes');
	fn(tx);
	return tx.commit();
};

describe('note operations', () => {
	it('sets, edits and removes notes, marking comments dirty, with undo', () => {
		const wb = createWorkbook();
		wb.sheets[0].commentsDirty = false;
		const entry = run(wb, (tx) => setNote(tx, wb, 0, 2, 1, 'Check this', 'Ann'));
		expect(noteAt(wb, 0, 2, 1)).toMatchObject({ text: 'Check this', author: 'Ann' });
		expect(wb.sheets[0].commentsDirty).toBe(true);
		run(wb, (tx) => setNote(tx, wb, 0, 2, 1, 'Edited', 'Bob'));
		expect(noteAt(wb, 0, 2, 1)).toMatchObject({ text: 'Edited', author: 'Ann' });
		run(wb, (tx) => setNote(tx, wb, 0, 2, 1, '  ', 'Bob'));
		expect(noteAt(wb, 0, 2, 1)).toBeUndefined();
		entry.undo();
		expect(wb.sheets[0].notes).toHaveLength(0);
	});

	it('shows, hides, deletes and navigates notes', () => {
		const wb = createWorkbook();
		run(wb, (tx) => {
			setNote(tx, wb, 0, 5, 0, 'b', 'x');
			setNote(tx, wb, 0, 1, 3, 'a', 'x');
			setNote(tx, wb, 0, 5, 2, 'c', 'x');
		});
		run(wb, (tx) => setNoteVisible(tx, wb, 0, 1, 3, true));
		expect(noteAt(wb, 0, 1, 3)?.visible).toBe(true);
		run(wb, (tx) => setAllNotesVisible(tx, wb, 0, true));
		expect(wb.sheets[0].notes.every((n) => n.visible)).toBe(true);
		expect(adjacentNote(wb, 0, { row: 1, col: 3 }, 1)?.text).toBe('b');
		expect(adjacentNote(wb, 0, { row: 5, col: 2 }, 1)?.text).toBe('a');
		expect(adjacentNote(wb, 0, { row: 1, col: 3 }, -1)?.text).toBe('c');
		let removed = 0;
		run(wb, (tx) => {
			removed = deleteNotes(tx, wb, 0, [{ top: 5, left: 0, bottom: 5, right: 5 }]);
		});
		expect(removed).toBe(2);
		expect(wb.sheets[0].notes.map((n) => n.text)).toEqual(['a']);
	});
});

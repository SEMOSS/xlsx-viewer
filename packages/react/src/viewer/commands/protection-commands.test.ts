import {
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	protectSheet,
	setCellInput,
} from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStore } from '../state/viewer-store';
import { beginEdit, clearSelectionContents, commitEdit, updateEditText } from './edit-commands';
import { protectSheetCmd, protectWorkbookCmd, unprotectCmd } from './protection-commands';
import { insertSheetCmd } from './sheet-commands';

function protectedStore() {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	setCellInput(tx, wb, 0, 0, 0, 'locked');
	applyStylePatch(tx, wb, 0, [{ top: 1, left: 0, bottom: 1, right: 0 }], {
		protection: { locked: false },
	});
	protectSheet(tx, wb, 0, { sheet: true });
	tx.commit();
	return new ViewerStore(wb);
}

const select = (store: ViewerStore, row: number, col: number) =>
	store.setSelection({
		active: { row, col },
		anchor: { row, col },
		ranges: [{ top: row, left: col, bottom: row, right: col }],
	});

describe('protection in the viewer', () => {
	it('refuses to edit locked cells with an alert', () => {
		const store = protectedStore();
		beginEdit(store, { initialText: 'x' });
		expect(store.state.edit).toBeNull();
		expect(store.state.dialog?.kind).toBe('alert');
	});

	it('edits unlocked cells normally', () => {
		const store = protectedStore();
		select(store, 1, 0);
		beginEdit(store, { initialText: '4' });
		updateEditText(store, '42', 2);
		commitEdit(store, { move: 'none' });
		expect(store.activeSheet.cells.get(16384)?.value).toBe(42);
	});

	it('turns blocked transactions into an alert without changes', () => {
		const store = protectedStore();
		clearSelectionContents(store);
		expect(store.activeSheet.cells.get(0)?.value).toBe('locked');
		expect(store.state.dialog?.kind).toBe('alert');
		expect(store.isDirty).toBe(false);
	});

	it('protects and unprotects with a password', async () => {
		const store = new ViewerStore(createWorkbook());
		await protectSheetCmd(store, 'pw', { formatCells: true });
		expect(store.activeSheet.protection).toMatchObject({ sheet: true, algorithmName: 'SHA-512' });
		expect(await unprotectCmd(store, 'sheet', 'wrong')).toBe(false);
		expect(await unprotectCmd(store, 'sheet', 'pw')).toBe(true);
		expect(store.activeSheet.protection).toBeUndefined();
	}, 20000);

	it('blocks new sheets while the workbook structure is protected', async () => {
		const store = new ViewerStore(createWorkbook());
		await protectWorkbookCmd(store, '');
		insertSheetCmd(store);
		expect(store.workbook.sheets).toHaveLength(1);
		expect(store.state.dialog?.kind).toBe('alert');
	});
});

describe('protection and automatic row heights', () => {
	it('lets wrapped text in an unlocked cell grow its row on a protected sheet', () => {
		const store = protectedStore();
		store.setSelection({
			active: { row: 1, col: 0 },
			anchor: { row: 1, col: 0 },
			ranges: [{ top: 1, left: 0, bottom: 1, right: 0 }],
		});
		beginEdit(store, { initialText: 'a' });
		updateEditText(store, 'one\ntwo\nthree', 13);
		expect(commitEdit(store, { move: 'none' })).toBe(true);
		expect(store.activeSheet.cells.get(16384)?.value).toBe('one\ntwo\nthree');
		expect(store.activeSheet.rows.get(1)?.height).toBeGreaterThan(30);
		expect(store.state.dialog).toBeNull();
	});
});

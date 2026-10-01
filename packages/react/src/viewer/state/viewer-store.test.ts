import { cellKey, createWorkbook, deleteRows, setCellInput } from 'xlsx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ViewerStore } from './viewer-store';

const value = (store: ViewerStore, row: number, col: number) =>
	store.activeSheet.cells.get(cellKey(row, col))?.value;

describe('ViewerStore', () => {
	it('commits, undoes and redoes transactions and notifies subscribers', () => {
		const store = new ViewerStore(createWorkbook());
		const listener = vi.fn();
		store.subscribe(listener);
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 0, 0, '5'));
		expect(value(store, 0, 0)).toBe(5);
		expect(store.isDirty).toBe(true);
		expect(listener).toHaveBeenCalled();
		expect(store.getSnapshot().history.undo).toBe('Type');
		store.undo();
		expect(value(store, 0, 0)).toBeUndefined();
		expect(store.getSnapshot().history.redo).toBe('Type');
		store.redo();
		expect(value(store, 0, 0)).toBe(5);
	});

	it('rolls back a transaction that throws, without marking the workbook dirty', () => {
		const store = new ViewerStore(createWorkbook());
		expect(() =>
			store.transact('Broken', (tx) => {
				setCellInput(tx, store.workbook, 0, 1, 1, 'partial');
				throw new Error('boom');
			}),
		).toThrow('boom');
		expect(value(store, 1, 1)).toBeUndefined();
		expect(store.isDirty).toBe(false);
		expect(store.getSnapshot().history.undo).toBeNull();
	});

	it('refuses edits when read-only', () => {
		const store = new ViewerStore(createWorkbook(), { readOnly: true });
		expect(store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 0, 0, 'x'))).toBe(
			false,
		);
		expect(value(store, 0, 0)).toBeUndefined();
	});

	it('keeps a selection per sheet', () => {
		const store = new ViewerStore(createWorkbook({ sheetNames: ['A', 'B'] }));
		store.setSelection(
			{
				active: { row: 3, col: 2 },
				anchor: { row: 3, col: 2 },
				ranges: [{ top: 3, left: 2, bottom: 3, right: 2 }],
			},
			1,
		);
		expect(store.selectionOf(0).active).toEqual({ row: 0, col: 0 });
		expect(store.selectionOf(1).active).toEqual({ row: 3, col: 2 });
	});

	it('grows the used extent with edits and recomputes it after structural changes', () => {
		const store = new ViewerStore(createWorkbook());
		const sheet = store.activeSheet;
		expect(store.usedExtentOf(sheet)).toEqual({ rows: 0, cols: 0 });
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 9, 4, 'far'));
		expect(store.usedExtentOf(sheet)).toEqual({ rows: 10, cols: 5 });
		store.transact('Delete Rows', (tx) => deleteRows(tx, store.workbook, 0, 0, 5));
		expect(store.usedExtentOf(sheet)).toEqual({ rows: 5, cols: 5 });
	});
});

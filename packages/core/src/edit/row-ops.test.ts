import { describe, expect, it } from 'vitest';

import { createWorkbook } from '../sheet/workbook';
import { setAutoRowHeight, setRowHeight, setRowsHidden } from './layout-ops';
import { setCellInput } from './cell-input';
import { EditTransaction } from './transaction';

describe('row-level undo tracking', () => {
	it('undoes and redoes row heights without a sheet snapshot', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.rows.set(3, { height: 30, customHeight: true });
		const tx = new EditTransaction(wb, 'Row Height');
		setRowHeight(tx, wb, 0, 2, 3, 20);
		const entry = tx.commit();
		expect(sheet.rows.get(2)).toEqual({ height: 20, customHeight: true, hidden: undefined });
		entry.undo();
		expect(sheet.rows.has(2)).toBe(false);
		expect(sheet.rows.get(3)).toEqual({ height: 30, customHeight: true });
		entry.redo();
		expect(sheet.rows.get(3)?.height).toBe(20);
		expect(entry.structural).toBe(false);
	});

	it('stores automatic heights without the custom flag', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Auto');
		setAutoRowHeight(tx, wb, 0, 5, 45);
		tx.commit();
		expect(wb.sheets[0].rows.get(5)).toEqual({ height: 45, customHeight: undefined });
		const tx2 = new EditTransaction(wb, 'Auto');
		setAutoRowHeight(tx2, wb, 0, 5, null);
		tx2.commit();
		expect(wb.sheets[0].rows.has(5)).toBe(false);
	});

	it('restores rows touched before a later sheet snapshot', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		const tx = new EditTransaction(wb, 'Mixed');
		setRowsHidden(tx, wb, 0, 1, 1, true);
		tx.touchSheet(0);
		const entry = tx.commit();
		expect(sheet.rows.get(1)?.hidden).toBe(true);
		entry.undo();
		expect(sheet.rows.has(1)).toBe(false);
	});

	it('lists the cells touched one by one', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Typing');
		setCellInput(tx, wb, 0, 4, 2, 'hello');
		expect([...tx.touchedCellKeys(0)]).toHaveLength(1);
		expect(tx.touchedSheetIndices()).toEqual([0]);
	});
});

describe('case-sensitive sort order', () => {
	it('puts lowercase before uppercase only among otherwise equal texts', async () => {
		const { compareForSort } = await import('./sort-ops');
		const words = ['banana', 'APPLE', 'Apple', 'apple', 'Banana'];
		const sorted = [...words].sort((a, b) => compareForSort(a, b, false, true));
		expect(sorted).toEqual(['apple', 'Apple', 'APPLE', 'banana', 'Banana']);
		const insensitive = [...words].sort((a, b) => compareForSort(a, b, false, false));
		expect(insensitive.map((w) => w.toLowerCase())).toEqual([
			'apple',
			'apple',
			'apple',
			'banana',
			'banana',
		]);
	});
});

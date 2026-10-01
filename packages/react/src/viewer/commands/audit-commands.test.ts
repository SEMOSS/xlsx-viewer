import { createWorkbook, EditTransaction, setCellInput } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStore } from '../state/viewer-store';
import { removeArrows, traceDependents, tracePrecedents } from './audit-commands';

function setup() {
	const wb = createWorkbook({ sheetNames: ['S1', 'S2'] });
	const tx = new EditTransaction(wb, 'seed');
	setCellInput(tx, wb, 0, 0, 0, '1');
	setCellInput(tx, wb, 0, 1, 0, '2');
	setCellInput(tx, wb, 0, 2, 0, '=SUM(A1:A2)');
	setCellInput(tx, wb, 0, 0, 1, '=A3*2');
	setCellInput(tx, wb, 1, 0, 0, '=S1!A3+1');
	tx.commit();
	const store = new ViewerStore(wb);
	const select = (row: number, col: number) =>
		store.setSelection({
			active: { row, col },
			anchor: { row, col },
			ranges: [{ top: row, left: col, bottom: row, right: col }],
		});
	return { store, select };
}

describe('formula auditing', () => {
	it('traces precedents of a formula', () => {
		const { store, select } = setup();
		select(2, 0);
		tracePrecedents(store);
		expect(store.state.traceArrows).toEqual([
			{
				sheet: 0,
				from: { top: 0, left: 0, bottom: 1, right: 0 },
				to: { row: 2, col: 0 },
				kind: 'precedent',
			},
		]);
	});

	it('traces dependents, including other sheets, and removes arrows', () => {
		const { store, select } = setup();
		select(2, 0);
		traceDependents(store);
		const arrows = store.state.traceArrows;
		expect(arrows.some((a) => a.to.row === 0 && a.to.col === 1 && !a.external)).toBe(true);
		expect(arrows.some((a) => a.external)).toBe(true);
		removeArrows(store, 'dependent');
		expect(store.state.traceArrows).toHaveLength(0);
	});

	it('clears arrows when the workbook changes', () => {
		const { store, select } = setup();
		select(2, 0);
		tracePrecedents(store);
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 5, 5, 'x'));
		expect(store.state.traceArrows).toHaveLength(0);
	});
});

import { applyStylePatch, createWorkbook, setCellInput, setRowHeight } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStore } from '../../viewer/state/viewer-store';
import { lineHeightPt } from './row-height';

const A1 = { top: 0, left: 0, bottom: 0, right: 0 };

describe('lineHeightPt', () => {
	it('matches Excel for the default 11 pt font', () => {
		expect(lineHeightPt(11)).toBe(15);
		expect(lineHeightPt(12)).toBe(15.75);
		expect(lineHeightPt(24)).toBe(31.5);
	});
});

describe('automatic row height', () => {
	it('grows a row for a larger font and undoes with the edit', () => {
		const store = new ViewerStore(createWorkbook());
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 0, 0, 'Title'));
		expect(store.activeSheet.rows.get(0)?.height).toBeUndefined();
		store.transact('Font Size', (tx) =>
			applyStylePatch(tx, store.workbook, 0, [A1], { font: { size: 24 } }),
		);
		expect(store.activeSheet.rows.get(0)).toMatchObject({ height: 31.5 });
		expect(store.activeSheet.rows.get(0)?.customHeight).toBeFalsy();
		store.undo();
		expect(store.activeSheet.rows.get(0)?.height).toBeUndefined();
	});

	it('fits wrapped text and shrinks back when it is cleared', () => {
		const store = new ViewerStore(createWorkbook());
		store.transact('Type', (tx) => {
			setCellInput(tx, store.workbook, 0, 2, 0, 'one\ntwo\nthree');
		});
		expect(store.activeSheet.rows.get(2)?.height).toBe(45);
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 2, 0, 'short'));
		expect(store.activeSheet.rows.has(2)).toBe(false);
	});

	it('leaves rows with a custom height alone', () => {
		const store = new ViewerStore(createWorkbook());
		store.transact('Height', (tx) => setRowHeight(tx, store.workbook, 0, 1, 1, 9));
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 1, 0, 'a\nb\nc'));
		expect(store.activeSheet.rows.get(1)).toMatchObject({ height: 9, customHeight: true });
	});

	it('can be turned off per transaction', () => {
		const store = new ViewerStore(createWorkbook());
		store.transact('Type', (tx) => setCellInput(tx, store.workbook, 0, 0, 0, 'a\nb'), {
			autoRowHeight: false,
		});
		expect(store.activeSheet.rows.has(0)).toBe(false);
	});
});

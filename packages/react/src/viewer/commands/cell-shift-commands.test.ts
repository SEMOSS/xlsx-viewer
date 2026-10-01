import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import {
	cellKey,
	createWorkbook,
	EditTransaction,
	getCell,
	MAX_ROW_INDEX,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';

import { ViewerStore } from '../state/viewer-store';
import { deleteCellsShiftCmd, insertCellsShiftCmd, wholeAxisOfRanges } from './cell-shift-commands';

function storeWith(entries: Record<string, string>, readOnly = false) {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		const r = parseRange(ref)!;
		setCellInput(tx, wb, 0, r.top, r.left, text);
	}
	tx.commit();
	return new ViewerStore(wb, { readOnly });
}

function select(store: ViewerStore, ...refs: string[]) {
	const ranges = refs.map((r) => parseRange(r)!);
	const active = { row: ranges[ranges.length - 1].top, col: ranges[ranges.length - 1].left };
	store.setSelection({ active, anchor: active, ranges });
}

const value = (wb: Workbook, ref: string) => {
	const r = parseRange(ref)!;
	return getCell(wb.sheets[0], r.top, r.left)?.value ?? null;
};

describe('insertCellsShiftCmd / deleteCellsShiftCmd', () => {
	it('shifts every selected range in one undoable step', () => {
		const store = storeWith({ A2: 'a2', A5: 'a5', C2: 'c2', D1: '=A5+C2' });
		const wb = store.workbook;
		select(store, 'A2', 'C2');
		expect(insertCellsShiftCmd(store, 'down')).toBe(true);
		expect([value(wb, 'A3'), value(wb, 'A6'), value(wb, 'C3')]).toEqual(['a2', 'a5', 'c2']);
		expect(getCell(wb.sheets[0], 0, 3)?.formula).toBe('A6+C3');
		expect(store.getSnapshot().history.undo).toBe('Insert Cells');
		store.undo();
		expect([value(wb, 'A2'), value(wb, 'C2')]).toEqual(['a2', 'c2']);
		expect(getCell(wb.sheets[0], 0, 3)?.formula).toBe('A5+C2');
		store.redo();
		expect(value(wb, 'A3')).toBe('a2');
	});

	it('processes lower ranges first in the same columns', () => {
		const store = storeWith({ A1: '1', A2: '2', A3: '3', A4: '4' });
		select(store, 'A1', 'A3');
		expect(deleteCellsShiftCmd(store, 'up')).toBe(true);
		expect([
			value(store.workbook, 'A1'),
			value(store.workbook, 'A2'),
			value(store.workbook, 'A3'),
		]).toEqual([2, 4, null]);
		expect(store.getSnapshot().history.undo).toBe('Delete Cells');
	});

	it('shows Excel refusals and changes nothing', () => {
		const store = storeWith({ B5: 'x' });
		store.workbook.sheets[0].merges.push(parseRange('A9:C9')!);
		select(store, 'B3');
		expect(insertCellsShiftCmd(store, 'down')).toBe(false);
		expect(store.state.message).toBe("We can't do that to a merged cell.");
		expect(value(store.workbook, 'B5')).toBe('x');
		expect(store.isDirty).toBe(false);
		expect(store.getSnapshot().history.undo).toBeNull();
	});

	it('refuses overlapping selections', () => {
		const store = storeWith({ A1: 'x' });
		select(store, 'A1:B2', 'B2:C3');
		expect(deleteCellsShiftCmd(store, 'left')).toBe(false);
		expect(store.state.message).toBe("This action won't work on overlapping selections.");
		expect(value(store.workbook, 'A1')).toBe('x');
	});

	it('finds refusals caused by an earlier range and rolls the trial back', () => {
		const store = storeWith({ A1: 'top' });
		const sheet = store.workbook.sheets[0];
		// A value one row above the bottom: each range alone fits, both together push it off.
		sheet.cells.set(cellKey(MAX_ROW_INDEX - 1, 0), { value: 'near the end' });
		sheet.dirty = false;
		select(store, 'A1', 'A3');
		expect(insertCellsShiftCmd(store, 'down')).toBe(false);
		expect(store.state.message).toContain('cannot shift nonblank cells off of the worksheet');
		expect(sheet.cells.get(cellKey(MAX_ROW_INDEX - 1, 0))?.value).toBe('near the end');
		expect(value(store.workbook, 'A1')).toBe('top');
		expect(sheet.dirty).toBe(false);
		expect(store.isDirty).toBe(false);
	});

	it('inserts / deletes entire columns or rows for whole-axis selections, like Excel', () => {
		const store = storeWith({ A1: 'a', B1: 'b', B2: 'b2', D1: '=B2' });
		const wb = store.workbook;
		select(store, 'B:B');
		expect(wholeAxisOfRanges(store.selectionOf().ranges)).toBe('cols');
		expect(insertCellsShiftCmd(store, 'down')).toBe(true);
		expect([value(wb, 'B1'), value(wb, 'C1'), value(wb, 'C2')]).toEqual([null, 'b', 'b2']);
		expect(getCell(wb.sheets[0], 0, 4)?.formula).toBe('C2');
		expect(store.getSnapshot().history.undo).toBe('Insert Columns');
		select(store, '2:2');
		expect(deleteCellsShiftCmd(store, 'left')).toBe(true);
		expect(value(wb, 'C2')).toBeNull();
		expect(store.getSnapshot().history.undo).toBe('Delete Rows');
		expect(wholeAxisOfRanges([parseRange('A1:B2')!])).toBeNull();
	});

	it('does nothing when read-only', () => {
		const store = storeWith({ A1: 'x' }, true);
		select(store, 'A1');
		expect(insertCellsShiftCmd(store, 'right')).toBe(false);
		expect(value(store.workbook, 'A1')).toBe('x');
	});
});

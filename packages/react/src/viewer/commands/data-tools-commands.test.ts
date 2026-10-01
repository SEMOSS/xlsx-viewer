import { describe, expect, it } from 'vitest';
import {
	createWorkbook,
	EditTransaction,
	getCell,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';

import { createTranslator } from '../../i18n';
import { ViewerStore } from '../state/viewer-store';
import {
	removeDuplicatesCmd,
	removeDuplicatesMessage,
	textToColumnsCmd,
	textToColumnsResultRange,
} from './data-tools-commands';

function storeWith(column: string[], readOnly = false) {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	column.forEach((text, row) => setCellInput(tx, wb, 0, row, 0, text));
	tx.commit();
	return new ViewerStore(wb, { readOnly });
}

const r = (ref: string) => parseRange(ref)!;
const NO_DUPLICATES = 'No duplicate values found.';

describe('removeDuplicatesMessage', () => {
	it('uses Excel wording', () => {
		const t = createTranslator();
		expect(removeDuplicatesMessage({ removed: 3, remaining: 7 }, t)).toBe(
			'3 duplicate values found and removed; 7 unique values remain.',
		);
		expect(removeDuplicatesMessage({ removed: 1, remaining: 1 }, t)).toBe(
			'1 duplicate value found and removed; 1 unique value remains.',
		);
		expect(removeDuplicatesMessage({ removed: 0, remaining: 4 }, t)).toBe(NO_DUPLICATES);
	});
});

describe('removeDuplicatesCmd', () => {
	it('removes duplicates, selects the rest and reports', () => {
		const store = storeWith(['a', 'b', 'a', 'c', 'b']);
		const message = removeDuplicatesCmd(store, r('A1:A5'), { columns: [0], hasHeader: false });
		expect(message).toBe('2 duplicate values found and removed; 3 unique values remain.');
		expect(store.state.message).toBe(message);
		expect(store.selectionOf().ranges).toEqual([r('A1:A3')]);
		store.undo();
		expect(getCell(store.workbook.sheets[0], 4, 0)?.value).toBe('b');
		store.redo();
		expect(store.selectionOf().ranges).toEqual([r('A1:A3')]);
	});

	it('does nothing without duplicates or when read-only', () => {
		const store = storeWith(['a', 'b']);
		expect(removeDuplicatesCmd(store, r('A1:A2'), { columns: [0], hasHeader: false })).toBe(
			NO_DUPLICATES,
		);
		expect(store.isDirty).toBe(false);
		const ro = storeWith(['a', 'a'], true);
		expect(removeDuplicatesCmd(ro, r('A1:A2'), { columns: [0], hasHeader: false })).toBeNull();
		expect(getCell(ro.workbook.sheets[0], 1, 0)?.value).toBe('a');
	});
});

describe('textToColumnsCmd', () => {
	it('converts in one step and selects the result', () => {
		const store = storeWith(['a,b,c', 'd']);
		const options = { mode: 'delimited' as const, comma: true, destination: { row: 0, col: 0 } };
		expect(textToColumnsResultRange(store, r('A1:A2'), options)).toEqual(r('A1:C2'));
		expect(textToColumnsCmd(store, r('A1:A2'), options)).toBeNull();
		expect(getCell(store.workbook.sheets[0], 0, 2)?.value).toBe('c');
		expect(store.selectionOf().ranges).toEqual([r('A1:C2')]);
		expect(store.getSnapshot().history.undo).toBe('Text to Columns');
	});

	it('changes nothing when every column is skipped', () => {
		const store = storeWith(['a,b']);
		const options = {
			mode: 'delimited' as const,
			comma: true,
			destination: { row: 0, col: 0 },
			columnFormats: ['skip' as const, 'skip' as const],
		};
		expect(textToColumnsCmd(store, r('A1'), options)).toBeNull();
		expect(getCell(store.workbook.sheets[0], 0, 0)?.value).toBe('a,b');
		expect(store.isDirty).toBe(false);
	});

	it('returns refusals without changing anything', () => {
		const store = storeWith(['a,b']);
		store.workbook.sheets[0].merges.push(r('B1:C1'));
		const options = { mode: 'delimited' as const, comma: true, destination: { row: 0, col: 0 } };
		expect(textToColumnsCmd(store, r('A1'), options)).toBe("We can't do that to a merged cell.");
		expect(getCell(store.workbook.sheets[0], 0, 0)?.value).toBe('a,b');
		expect(store.isDirty).toBe(false);
	});
});

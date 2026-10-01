import { describe, expect, it } from 'vitest';
import { createWorkbook, EditTransaction, parseRange, setCellInput } from 'xlsx-viewer-core';

import { ViewerStore } from '../state/viewer-store';
import {
	defineNameCmd,
	definedNameKey,
	deleteNameCmd,
	editNameCmd,
	editNameDraft,
	nameManagerRows,
	newNameDraft,
} from './name-commands';

function makeStore(readOnly = false): ViewerStore {
	const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet'] });
	const tx = new EditTransaction(wb, 'seed');
	setCellInput(tx, wb, 0, 0, 1, 'Unit Price');
	setCellInput(tx, wb, 0, 1, 1, '12');
	setCellInput(tx, wb, 0, 2, 1, '5%');
	tx.commit();
	wb.definedNames.push(
		{ name: 'Price', formula: 'Data!$B$2', comment: 'Per unit' },
		{ name: '_xlnm.Print_Area', formula: "'My Sheet'!$A$1:$C$3", localSheet: 1, builtin: true },
		{
			name: '_xlnm._FilterDatabase',
			formula: 'Data!$A$1:$B$3',
			localSheet: 0,
			hidden: true,
			builtin: true,
		},
		{ name: 'base', formula: '100' },
	);
	return new ViewerStore(wb, { readOnly });
}

describe('name dialog helpers', () => {
	it('keys definitions by name and scope', () => {
		expect(definedNameKey({ name: 'Rate' })).toBe('rate|');
		expect(definedNameKey({ name: 'Rate', localSheet: 2 })).toBe('rate|2');
	});

	it('prefills New Name from the selection', () => {
		const store = makeStore();
		const range = parseRange('B2:B3')!;
		store.setSelection({ active: { row: 1, col: 1 }, anchor: { row: 1, col: 1 }, ranges: [range] });
		expect(newNameDraft(store)).toEqual({
			name: 'Unit_Price',
			scope: undefined,
			comment: '',
			refersTo: '=Data!$B$2:$B$3',
		});
		expect(editNameDraft(store.workbook.definedNames[1])).toEqual({
			name: 'Print_Area',
			scope: 1,
			comment: '',
			refersTo: "='My Sheet'!$A$1:$C$3",
		});
	});

	it('lists visible names sorted with their value, reference, scope and comment', () => {
		const rows = nameManagerRows(makeStore());
		expect(rows.map((r) => [r.name, r.value, r.refersTo, r.scope, r.comment])).toEqual([
			['base', '100', '=100', 'Workbook', ''],
			['Price', '12', '=Data!$B$2', 'Workbook', 'Per unit'],
			['Print_Area', '{"","","";"","","";"","",""}', "='My Sheet'!$A$1:$C$3", 'My Sheet', ''],
		]);
	});
});

describe('name commands', () => {
	it('defines, edits and deletes names as undoable steps', () => {
		const store = makeStore();
		const added = defineNameCmd(store, { name: 'Tax', formula: 'Data!$B$3' });
		expect(added).toEqual({ def: { name: 'Tax', formula: 'Data!$B$3' } });
		expect(store.getSnapshot().history.undo).toBe('Define Name');

		const edited = editNameCmd(store, added.def!, { name: 'VAT', formula: 'Data!$B$3' });
		expect(edited.def?.name).toBe('VAT');
		expect(store.getSnapshot().history.undo).toBe('Edit Name');

		expect(deleteNameCmd(store, { name: 'vat', formula: '' })).toBe(true);
		expect(store.getSnapshot().history.undo).toBe('Delete Name');
		expect(store.workbook.definedNames.some((n) => n.name === 'VAT')).toBe(false);

		store.undo();
		expect(store.workbook.definedNames.some((n) => n.name === 'VAT')).toBe(true);
		store.undo();
		store.undo();
		expect(store.workbook.definedNames.some((n) => /tax|vat/i.test(n.name))).toBe(false);
	});

	it('reports rejected names and skips unchanged edits', () => {
		const store = makeStore();
		expect(defineNameCmd(store, { name: 'price', formula: '1' }).error).toBe(
			'The name already exists. Enter a unique name.',
		);
		expect(store.getSnapshot().history.undo).toBeNull();
		const price = store.workbook.definedNames[0];
		expect(editNameCmd(store, price, { ...price, comment: ' Per unit ' })).toEqual({ def: price });
		expect(store.getSnapshot().history.undo).toBeNull();
	});

	it('does nothing when read-only', () => {
		const store = makeStore(true);
		expect(defineNameCmd(store, { name: 'Tax', formula: '1' })).toEqual({ def: undefined });
		expect(deleteNameCmd(store, { name: 'Price', formula: '' })).toBe(false);
		expect(store.workbook.definedNames).toHaveLength(4);
	});
});

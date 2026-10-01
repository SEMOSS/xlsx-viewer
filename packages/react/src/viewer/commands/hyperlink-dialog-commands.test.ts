import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import {
	createWorkbook,
	EditTransaction,
	getCell,
	mergeRange,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';

import { ViewerStore } from '../state/viewer-store';
import {
	applyHyperlinkForm,
	buildHyperlink,
	defaultLinkText,
	hyperlinkDialogSetup,
	type HyperlinkForm,
} from './hyperlink-dialog-commands';

function book(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet', 'Hidden'] });
	wb.sheets[2].state = 'hidden';
	const tx = new EditTransaction(wb, 'seed');
	setCellInput(tx, wb, 0, 0, 0, '1234.5');
	setCellInput(tx, wb, 0, 0, 1, 'Docs');
	setCellInput(tx, wb, 0, 1, 1, 'More');
	tx.commit();
	wb.definedNames.push(
		{ name: 'Totals', formula: 'Data!$C$1' },
		{ name: 'Local', formula: 'Data!$C$2', localSheet: 0 },
		{ name: 'Elsewhere', formula: "'My Sheet'!$A$1", localSheet: 1 },
		{ name: 'Secret', formula: '1', hidden: true },
		{ name: '_xlnm.Print_Area', formula: 'Data!$A$1:$B$2', localSheet: 0, builtin: true },
	);
	return wb;
}

function storeAt(ref: string, wb = book()): ViewerStore {
	const store = new ViewerStore(wb);
	const range = parseRange(ref)!;
	const active = { row: range.top, col: range.left };
	store.setSelection({ active, anchor: active, ranges: [range] });
	return store;
}

const form = (patch: Partial<HyperlinkForm>): HyperlinkForm => ({
	kind: 'file',
	text: '',
	tooltip: '',
	address: '',
	cellRef: 'A1',
	place: 'sheet:0',
	email: '',
	subject: '',
	...patch,
});

describe('hyperlinkDialogSetup', () => {
	it('starts an empty cell on Existing File or Web Page with the text following the address', () => {
		const setup = hyperlinkDialogSetup(storeAt('D4'));
		expect(setup).toMatchObject({
			editing: false,
			sheetIndex: 0,
			mirror: true,
			range: parseRange('D4'),
		});
		expect(setup.form).toEqual(form({}));
		expect(setup.places.map((p) => `${p.group}:${p.label}`)).toEqual([
			'sheet:Data',
			'sheet:My Sheet',
			'name:Local',
			'name:Totals',
		]);
	});

	it('shows the displayed text of the cell, or nothing for several values', () => {
		expect(hyperlinkDialogSetup(storeAt('A1')).form.text).toBe('1234.5');
		expect(hyperlinkDialogSetup(storeAt('A1')).mirror).toBe(false);
		expect(hyperlinkDialogSetup(storeAt('A1:A3')).form.text).toBe('1234.5');
		expect(hyperlinkDialogSetup(storeAt('A1:B2')).form.text).toBeNull();
	});

	it('expands the range to merged areas', () => {
		const wb = book();
		const tx = new EditTransaction(wb, 'merge');
		mergeRange(tx, wb, 0, parseRange('D1:E2')!);
		tx.commit();
		expect(hyperlinkDialogSetup(storeAt('E1', wb)).range).toEqual(parseRange('D1:E2'));
	});

	it('reads an existing link back into the form', () => {
		const wb = book();
		const links = wb.sheets[0].hyperlinks;
		links.push(
			{ ref: parseRange('C1')!, target: 'https://example.com/a', location: 'top', tooltip: 'Tip' },
			{ ref: parseRange('C2')!, target: 'mailto:a@b.com?subject=Hi%20there' },
			{ ref: parseRange('C3')!, location: "'My Sheet'!B4" },
			{ ref: parseRange('C4')!, location: 'totals' },
			{ ref: parseRange('C5')!, location: 'E7' },
		);
		const at = (ref: string) => hyperlinkDialogSetup(storeAt(ref, wb));
		expect(at('C1')).toMatchObject({ editing: true });
		expect(at('C1').form).toMatchObject({
			kind: 'file',
			address: 'https://example.com/a#top',
			tooltip: 'Tip',
		});
		expect(at('C2').form).toMatchObject({ kind: 'email', email: 'a@b.com', subject: 'Hi there' });
		expect(at('C3').form).toMatchObject({ kind: 'place', place: 'sheet:1', cellRef: 'B4' });
		expect(at('C4').form).toMatchObject({ kind: 'place', place: 'name:totals|' });
		expect(at('C5').form).toMatchObject({ kind: 'place', place: 'sheet:0', cellRef: 'E7' });
	});
});

describe('buildHyperlink', () => {
	const places = hyperlinkDialogSetup(storeAt('A1')).places;

	it('completes web and e-mail addresses and keeps internal ones', () => {
		expect(buildHyperlink(form({ address: ' www.example.com ' }), places)).toEqual({
			link: { target: 'http://www.example.com' },
			display: 'www.example.com',
		});
		expect(buildHyperlink(form({ address: '#Data!B2', text: 'Go' }), places)).toEqual({
			link: { location: 'Data!B2' },
			display: 'Go',
		});
		expect(
			buildHyperlink(
				form({ kind: 'email', email: 'a@b.com', subject: 'Q 1', tooltip: ' Mail ' }),
				places,
			),
		).toEqual({
			link: { target: 'mailto:a@b.com?subject=Q%201', tooltip: 'Mail' },
			display: 'mailto:a@b.com?subject=Q%201',
		});
		expect(buildHyperlink(form({ address: '   ' }), places)).toBeNull();
		expect(buildHyperlink(form({ kind: 'email' }), places)).toBeNull();
	});

	it('builds places in this document', () => {
		expect(
			buildHyperlink(form({ kind: 'place', place: 'sheet:1', cellRef: ' b2 ' }), places),
		).toEqual({
			link: { location: "'My Sheet'!B2" },
			display: "'My Sheet'!B2",
		});
		expect(
			buildHyperlink(form({ kind: 'place', place: 'name:totals|', text: null }), places),
		).toEqual({
			link: { location: 'Totals' },
		});
		expect(buildHyperlink(form({ kind: 'place', cellRef: 'ZZZZ1' }), places)).toEqual({
			error: 'message.invalidReference',
		});
	});

	it('derives the default text per link type', () => {
		expect(defaultLinkText(form({ kind: 'place', place: 'sheet:0', cellRef: 'c3' }), places)).toBe(
			'Data!C3',
		);
		expect(defaultLinkText(form({ kind: 'place', place: 'name:local|0' }), places)).toBe('Local');
		expect(defaultLinkText(form({ kind: 'email', email: '' }), places)).toBe('');
	});
});

describe('applyHyperlinkForm', () => {
	it('inserts and edits links as undoable steps', () => {
		const store = storeAt('D4');
		const setup = hyperlinkDialogSetup(store);
		expect(
			applyHyperlinkForm(store, setup, form({ address: 'www.example.com', text: 'Site' })),
		).toBeNull();
		expect(store.getSnapshot().history.undo).toBe('Insert Hyperlink');
		expect(store.activeSheet.hyperlinks).toEqual([
			{ ref: parseRange('D4'), target: 'http://www.example.com', display: 'Site' },
		]);
		expect(getCell(store.activeSheet, 3, 3)?.value).toBe('Site');

		const again = hyperlinkDialogSetup(store);
		expect(again.editing).toBe(true);
		expect(
			applyHyperlinkForm(store, again, { ...again.form, address: 'https://other.example' }),
		).toBeNull();
		expect(store.getSnapshot().history.undo).toBe('Edit Hyperlink');
		expect(store.activeSheet.hyperlinks.map((h) => h.target)).toEqual(['https://other.example']);
		store.undo();
		expect(store.activeSheet.hyperlinks.map((h) => h.target)).toEqual(['http://www.example.com']);
	});

	it('reports a bad reference without changing anything', () => {
		const store = storeAt('D4');
		const setup = hyperlinkDialogSetup(store);
		expect(applyHyperlinkForm(store, setup, form({ kind: 'place', cellRef: 'nope' }))).toBe(
			"Reference isn't valid.",
		);
		expect(store.getSnapshot().history.undo).toBeNull();
		expect(applyHyperlinkForm(store, setup, form({}))).toBeNull();
		expect(store.isDirty).toBe(false);
	});
});

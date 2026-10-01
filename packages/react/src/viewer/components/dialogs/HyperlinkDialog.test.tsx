import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, findHyperlinkCellStyle, getCell, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { HyperlinkDialog } from './HyperlinkDialog';

function storeAt(ref: string, setup?: (store: ViewerStore) => void, readOnly = false) {
	const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet'] });
	typeCells(wb, { A1: 'Home', B1: '12', B2: 'x' });
	wb.definedNames.push({ name: 'Totals', formula: 'Data!$C$1' });
	const store = new ViewerStore(wb, { readOnly });
	setup?.(store);
	selectRange(store, ref);
	return store;
}

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;
const tab = (name: string) => screen.getByRole('tab', { name });
const ok = () => fireEvent.click(screen.getByRole('button', { name: 'OK' }));
const cell = (store: ViewerStore, a1: string) => {
	const r = parseRange(a1)!;
	return getCell(store.activeSheet, r.top, r.left);
};

describe('HyperlinkDialog', () => {
	it('links an empty cell to a web page, mirroring the address as the text', () => {
		const store = storeAt('D4');
		const { onClose } = renderDialog(HyperlinkDialog, { store });
		expect(screen.getByRole('dialog', { name: 'Insert Hyperlink' })).toBeTruthy();
		const tablist = screen.getByRole('tablist', { name: 'Link to:' });
		expect(tablist.getAttribute('aria-orientation')).toBe('vertical');
		expect(
			within(tablist)
				.getAllByRole('tab')
				.map((t) => t.textContent),
		).toEqual(['Existing File or Web Page', 'Place in This Document', 'E-mail Address']);
		expect(tab('Existing File or Web Page').getAttribute('aria-selected')).toBe('true');
		expect(document.activeElement).toBe(field('Address:'));
		expect(screen.queryByRole('button', { name: 'Remove Link' })).toBeNull();

		fireEvent.change(field('Address:'), { target: { value: 'www.example.com' } });
		expect(field('Text to display:').value).toBe('www.example.com');
		fireEvent.keyDown(field('Address:'), { key: 'Enter' });
		expect(onClose).toHaveBeenCalled();
		expect(store.activeSheet.hyperlinks).toEqual([
			{ ref: parseRange('D4'), target: 'http://www.example.com', display: 'www.example.com' },
		]);
		const linked = cell(store, 'D4')!;
		expect(linked.value).toBe('www.example.com');
		expect(store.workbook.styles.cellXfs[linked.style!].xfId).toBe(
			findHyperlinkCellStyle(store.workbook.styles),
		);
		expect(store.getSnapshot().history.undo).toBe('Insert Hyperlink');
	});

	it('keeps the cell text, adds a ScreenTip and switches tabs with the keyboard', () => {
		const store = storeAt('A1');
		renderDialog(HyperlinkDialog, { store });
		expect(field('Text to display:').value).toBe('Home');
		const tip = screen.getByRole('button', { name: 'ScreenTip...' });
		expect(tip.getAttribute('aria-expanded')).toBe('false');
		fireEvent.click(tip);
		fireEvent.change(field('ScreenTip text:'), { target: { value: 'Back to start' } });
		fireEvent.keyDown(tab('Existing File or Web Page'), { key: 'ArrowDown' });
		expect(tab('Place in This Document').getAttribute('aria-selected')).toBe('true');
		expect(document.activeElement).toBe(tab('Place in This Document'));
		fireEvent.keyDown(tab('Place in This Document'), { key: 'End' });
		expect(tab('E-mail Address').getAttribute('aria-selected')).toBe('true');
		fireEvent.change(field('E-mail address:'), { target: { value: 'team@example.com' } });
		fireEvent.change(field('Subject:'), { target: { value: 'Hello there' } });
		expect(field('Text to display:').value).toBe('Home');
		ok();
		expect(store.activeSheet.hyperlinks).toEqual([
			{
				ref: parseRange('A1'),
				target: 'mailto:team@example.com?subject=Hello%20there',
				tooltip: 'Back to start',
				display: 'Home',
			},
		]);
		expect(cell(store, 'A1')?.value).toBe('Home');
	});

	it('links to a place in this document and validates the reference', () => {
		const store = storeAt('D4');
		const { onClose } = renderDialog(HyperlinkDialog, { store });
		fireEvent.click(tab('Place in This Document'));
		expect(field('Type the cell reference:').value).toBe('A1');
		const list = screen.getByRole('listbox', { name: 'Or select a place in this document:' });
		const groups = within(list).getAllByRole('group');
		expect(
			groups.map((g) =>
				within(g)
					.getAllByRole('option')
					.map((o) => o.textContent),
			),
		).toEqual([['Data', 'My Sheet'], ['Totals']]);
		expect(within(list).getByRole('option', { name: 'Data' }).getAttribute('aria-selected')).toBe(
			'true',
		);
		fireEvent.keyDown(list, { key: 'ArrowDown' });
		expect(
			within(list).getByRole('option', { name: 'My Sheet' }).getAttribute('aria-selected'),
		).toBe('true');
		fireEvent.change(field('Type the cell reference:'), { target: { value: 'ZZZZ9' } });
		ok();
		expect(screen.getByRole('alert').textContent).toBe("Reference isn't valid.");
		expect(onClose).not.toHaveBeenCalled();
		fireEvent.change(field('Type the cell reference:'), { target: { value: 'c5' } });
		expect(field('Text to display:').value).toBe("'My Sheet'!C5");
		fireEvent.keyDown(list, { key: 'Enter' });
		expect(onClose).toHaveBeenCalled();
		expect(store.activeSheet.hyperlinks[0]).toEqual({
			ref: parseRange('D4'),
			location: "'My Sheet'!C5",
			display: "'My Sheet'!C5",
		});
	});

	it('links to a defined name, disabling the cell reference', () => {
		const store = storeAt('D4');
		renderDialog(HyperlinkDialog, { store });
		fireEvent.click(tab('Place in This Document'));
		fireEvent.click(screen.getByRole('option', { name: 'Totals' }));
		expect(field('Type the cell reference:').disabled).toBe(true);
		expect(field('Text to display:').value).toBe('Totals');
		ok();
		expect(store.activeSheet.hyperlinks[0]).toMatchObject({
			location: 'Totals',
			display: 'Totals',
		});
	});

	it('does not write text over a selection holding several values', () => {
		const store = storeAt('A1:B2');
		renderDialog(HyperlinkDialog, { store });
		expect(field('Text to display:').value).toBe('<<Selection in Document>>');
		expect(field('Text to display:').disabled).toBe(true);
		fireEvent.change(field('Address:'), { target: { value: 'https://example.com' } });
		ok();
		expect(store.activeSheet.hyperlinks).toEqual([
			{ ref: parseRange('A1:B2'), target: 'https://example.com' },
		]);
		expect(cell(store, 'A1')?.value).toBe('Home');
	});

	it('edits an existing link and removes it with Remove Link', () => {
		const link = (store: ViewerStore) =>
			store.activeSheet.hyperlinks.push({
				ref: parseRange('B2')!,
				target: 'https://old.example',
				tooltip: 'Old',
			});
		const store = storeAt('B2', link);
		const { onClose, unmount } = renderDialog(HyperlinkDialog, { store });
		expect(screen.getByRole('dialog', { name: 'Edit Hyperlink' })).toBeTruthy();
		expect(field('Address:').value).toBe('https://old.example');
		expect(field('ScreenTip text:').value).toBe('Old');
		fireEvent.change(field('Address:'), { target: { value: 'https://new.example' } });
		ok();
		expect(store.getSnapshot().history.undo).toBe('Edit Hyperlink');
		expect(store.activeSheet.hyperlinks).toEqual([
			{ ref: parseRange('B2'), target: 'https://new.example', tooltip: 'Old', display: 'x' },
		]);
		unmount();
		onClose.mockClear();

		const again = renderDialog(HyperlinkDialog, { store });
		fireEvent.click(screen.getByRole('button', { name: 'Remove Link' }));
		expect(again.onClose).toHaveBeenCalled();
		expect(store.activeSheet.hyperlinks).toEqual([]);
		expect(cell(store, 'B2')).toEqual({ value: 'x' });
	});

	it('cancels with Escape and disables OK and Remove Link when read-only', () => {
		const link = (store: ViewerStore) =>
			store.activeSheet.hyperlinks.push({ ref: parseRange('A1')!, location: 'Data!B1' });
		const store = storeAt('A1', link, true);
		const { onClose } = renderDialog(HyperlinkDialog, { store });
		expect(tab('Place in This Document').getAttribute('aria-selected')).toBe('true');
		expect(field('Type the cell reference:').value).toBe('B1');
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
		expect(
			(screen.getByRole('button', { name: 'Remove Link' }) as HTMLButtonElement).disabled,
		).toBe(true);
		fireEvent.keyDown(field('Type the cell reference:'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
	});
});

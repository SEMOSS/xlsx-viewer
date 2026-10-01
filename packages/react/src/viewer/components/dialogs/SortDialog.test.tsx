import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { cellKey, createWorkbook, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { SortDialog } from './SortDialog';

const SORT_EMPTY_MESSAGE =
	"This can't be applied to the selected range. Select a single cell in a range and try again.";

function peopleStore(select = 'B3', readOnly = false) {
	const wb = createWorkbook();
	typeCells(wb, {
		A1: 'Name',
		B1: 'Age',
		C1: 'Joined',
		A2: 'carol',
		B2: '41',
		C2: '3/1/2021',
		A3: 'Alice',
		B3: '29',
		C3: '1/15/2020',
		A4: 'bob',
		B4: '35',
		C4: '7/4/2019',
		A5: 'Alice',
		B5: '52',
		C5: '2/2/2022',
	});
	const store = new ViewerStore(wb, { readOnly });
	selectRange(store, select);
	return store;
}

const values = (store: ViewerStore, col: number, from = 1, to = 4) => {
	const out: unknown[] = [];
	for (let row = from; row <= to; row++) {
		out.push(store.activeSheet.cells.get(cellKey(row, col))?.value ?? null);
	}
	return out;
};

const select = (label: string) => screen.getByLabelText(label) as HTMLSelectElement;
const optionTexts = (label: string) => Array.from(select(label).options).map((o) => o.textContent);
const click = (name: string) => fireEvent.click(screen.getByRole('button', { name }));
const rows = () => screen.getAllByRole('row').slice(1);

describe('SortDialog', () => {
	it('selects the current region and sorts by the active cell column', () => {
		const store = peopleStore('B3');
		const { onClose } = renderDialog(SortDialog, { store });
		expect(store.selectionOf().ranges).toEqual([parseRange('A2:C5')]);
		expect((screen.getByLabelText('My data has headers') as HTMLInputElement).checked).toBe(true);
		expect(within(rows()[0]).getByRole('rowheader').textContent).toBe('Sort by');
		expect(select('Level 1 column').value).toBe('1');
		expect(optionTexts('Level 1 column')).toEqual(['Name', 'Age', 'Joined']);
		expect(optionTexts('Level 1 order')).toEqual(['Smallest to Largest', 'Largest to Smallest']);
		expect(select('Level 1 sort on').value).toBe('values');
		fireEvent.change(select('Level 1 order'), { target: { value: 'desc' } });
		click('OK');
		expect(onClose).toHaveBeenCalled();
		expect(values(store, 1)).toEqual([52, 41, 35, 29]);
		expect(values(store, 0, 0, 0)).toEqual(['Name']);
		expect(store.getSnapshot().history.undo).toBe('Sort');
	});

	it('words the order for text and date columns', () => {
		renderDialog(SortDialog, { store: peopleStore('A2') });
		expect(optionTexts('Level 1 order')).toEqual(['A to Z', 'Z to A']);
		fireEvent.change(select('Level 1 column'), { target: { value: '2' } });
		expect(optionTexts('Level 1 order')).toEqual(['Oldest to Newest', 'Newest to Oldest']);
	});

	it('adds, copies, moves and deletes levels with Excel checks on OK', () => {
		const store = peopleStore('B2');
		const { onClose } = renderDialog(SortDialog, { store });
		click('Add Level');
		expect(rows()).toHaveLength(2);
		expect(within(rows()[1]).getByRole('rowheader').textContent).toBe('Then by');
		expect(select('Level 2 column').value).toBe('');
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'All sort criteria must have a column specified. Check the selected sort criteria and try again.',
		);
		fireEvent.change(select('Level 2 column'), { target: { value: '0' } });
		expect(screen.queryByRole('alert')).toBeNull();
		click('Copy Level');
		expect(rows()).toHaveLength(3);
		expect(select('Level 3 column').value).toBe('0');
		expect(rows()[2].getAttribute('aria-selected')).toBe('true');
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'Name is being sorted by values more than once. Delete the duplicate sort condition and try again.',
		);
		click('Delete Level');
		expect(rows()).toHaveLength(2);
		expect(rows()[1].getAttribute('aria-selected')).toBe('true');
		click('Move Up');
		expect(select('Level 1 column').value).toBe('0');
		expect(select('Level 2 column').value).toBe('1');
		expect((screen.getByRole('button', { name: 'Move Up' }) as HTMLButtonElement).disabled).toBe(
			true,
		);
		fireEvent.change(select('Level 2 order'), { target: { value: 'desc' } });
		click('OK');
		expect(onClose).toHaveBeenCalled();
		expect(values(store, 0)).toEqual(['Alice', 'Alice', 'bob', 'carol']);
		expect(values(store, 1)).toEqual([52, 29, 35, 41]);
	});

	it('sorts the first row too without headers and labels columns by letter', () => {
		const store = peopleStore('B3');
		renderDialog(SortDialog, { store });
		fireEvent.click(screen.getByLabelText('My data has headers'));
		expect(store.selectionOf().ranges).toEqual([parseRange('A1:C5')]);
		expect(optionTexts('Level 1 column')).toEqual(['Column A', 'Column B', 'Column C']);
		expect(optionTexts('Level 1 order')).toEqual(['A to Z', 'Z to A']);
		click('OK');
		expect(values(store, 1, 0, 4)).toEqual([29, 35, 41, 52, 'Age']);
	});

	it('changes case sensitivity and orientation in Sort Options', () => {
		const store = peopleStore('B3');
		renderDialog(SortDialog, { store });
		click('Options...');
		const options = screen.getByRole('dialog', { name: 'Sort Options' });
		expect((within(options).getByLabelText('Sort top to bottom') as HTMLInputElement).checked).toBe(
			true,
		);
		fireEvent.keyDown(within(options).getByLabelText('Case sensitive'), { key: 'Escape' });
		expect(screen.queryByRole('dialog', { name: 'Sort Options' })).toBeNull();
		expect(screen.getByRole('dialog', { name: 'Sort' })).toBeTruthy();

		click('Options...');
		const again = screen.getByRole('dialog', { name: 'Sort Options' });
		fireEvent.click(within(again).getByLabelText('Case sensitive'));
		fireEvent.click(within(again).getByLabelText('Sort left to right'));
		fireEvent.click(within(again).getByRole('button', { name: 'OK' }));
		expect(screen.queryByRole('dialog', { name: 'Sort Options' })).toBeNull();
		expect(screen.getByRole('columnheader', { name: 'Row' })).toBeTruthy();
		expect(optionTexts('Level 1 row')).toEqual(['Row 1', 'Row 2', 'Row 3', 'Row 4', 'Row 5']);
		expect(select('Level 1 row').value).toBe('2');
		const headers = screen.getByLabelText('My data has headers') as HTMLInputElement;
		expect(headers.disabled).toBe(true);
		expect(headers.checked).toBe(false);
		expect(store.selectionOf().ranges).toEqual([parseRange('A1:C5')]);

		click('Options...');
		const third = screen.getByRole('dialog', { name: 'Sort Options' });
		expect((within(third).getByLabelText('Case sensitive') as HTMLInputElement).checked).toBe(true);
		expect((within(third).getByLabelText('Sort left to right') as HTMLInputElement).checked).toBe(
			true,
		);
	});

	it('sorts left to right by the chosen row', () => {
		const wb = createWorkbook();
		typeCells(wb, { A1: 'c', B1: 'a', C1: 'b', A2: 1, B2: 2, C2: 3 });
		const store = new ViewerStore(wb);
		selectRange(store, 'A1:C2');
		renderDialog(SortDialog, { store });
		click('Options...');
		fireEvent.click(screen.getByLabelText('Sort left to right'));
		fireEvent.keyDown(screen.getByLabelText('Sort left to right'), { key: 'Enter' });
		expect(select('Level 1 row').value).toBe('0');
		click('OK');
		expect([0, 1, 2].map((c) => values(store, c, 0, 1))).toEqual([
			['a', 2],
			['b', 3],
			['c', 1],
		]);
	});

	it('reports merged cells inline and keeps the dialog open', () => {
		const store = peopleStore('B3');
		store.workbook.sheets[0].merges.push(parseRange('A2:A3')!);
		const { onClose } = renderDialog(SortDialog, { store });
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'To do this, all the merged cells need to be the same size.',
		);
		expect(onClose).not.toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
	});

	it('closes with a message when there is nothing to sort', () => {
		const store = new ViewerStore(createWorkbook());
		const { onClose } = renderDialog(SortDialog, { store });
		expect(onClose).toHaveBeenCalled();
		expect(store.state.message).toBe(SORT_EMPTY_MESSAGE);
		expect(screen.queryByRole('dialog')).toBeNull();
	});

	it('limits the dialog to 64 levels', () => {
		renderDialog(SortDialog, { store: peopleStore() });
		const add = screen.getByRole('button', { name: 'Add Level' }) as HTMLButtonElement;
		for (let i = 0; i < 63; i++) {
			fireEvent.click(add);
		}
		expect(rows()).toHaveLength(64);
		expect(add.disabled).toBe(true);
		expect((screen.getByRole('button', { name: 'Copy Level' }) as HTMLButtonElement).disabled).toBe(
			true,
		);
	});

	it('cancels with Escape or Cancel without sorting', () => {
		const store = peopleStore('B3');
		const first = renderDialog(SortDialog, { store });
		fireEvent.change(select('Level 1 order'), { target: { value: 'desc' } });
		fireEvent.keyDown(select('Level 1 order'), { key: 'Escape' });
		expect(first.onClose).toHaveBeenCalled();
		first.unmount();
		const second = renderDialog(SortDialog, { store });
		click('Cancel');
		expect(second.onClose).toHaveBeenCalled();
		expect(values(store, 1)).toEqual([41, 29, 35, 52]);
		expect(store.isDirty).toBe(false);
	});

	it('disables OK on a read-only workbook', () => {
		const store = peopleStore('B3', true);
		renderDialog(SortDialog, { store });
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
		fireEvent.keyDown(select('Level 1 order'), { key: 'Enter' });
		expect(values(store, 1)).toEqual([41, 29, 35, 52]);
	});
});

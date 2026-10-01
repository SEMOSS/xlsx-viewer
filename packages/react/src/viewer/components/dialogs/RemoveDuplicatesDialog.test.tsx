import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import { createWorkbook, getCell, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { RemoveDuplicatesDialog } from './RemoveDuplicatesDialog';

function people(): Workbook {
	const wb = createWorkbook();
	typeCells(wb, {
		A1: 'Name',
		B1: 'Qty',
		A2: 'Ann',
		B2: '1',
		A3: 'Bob',
		B3: '2',
		A4: 'ann',
		B4: '3',
		A5: 'Bob',
		B5: '2',
		D1: 'apart',
	});
	return wb;
}

function open(wb: Workbook, ref = 'A2', readOnly = false) {
	const store = new ViewerStore(wb, { readOnly });
	selectRange(store, ref);
	return renderDialog(RemoveDuplicatesDialog, { store });
}

const column = (wb: Workbook, ref: string) => {
	const r = parseRange(ref)!;
	const out: unknown[] = [];
	for (let row = r.top; row <= r.bottom; row++) {
		out.push(getCell(wb.sheets[0], row, r.left)?.value ?? null);
	}
	return out;
};

const box = (name: string) => screen.getByRole('checkbox', { name }) as HTMLInputElement;
const okButton = () => screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement;

describe('RemoveDuplicatesDialog', () => {
	it('lists the columns of the current region with headers and selects the data rows', () => {
		const { store } = open(people());
		expect(screen.getByRole('dialog', { name: 'Remove Duplicates' })).toBeTruthy();
		expect(
			screen.getByText(
				'To delete duplicate values, select one or more columns that contain duplicates.',
			),
		).toBeTruthy();
		expect(box('My data has headers').checked).toBe(true);
		const list = screen.getByRole('group', { name: 'Columns' });
		expect([...list.querySelectorAll('label')].map((l) => l.textContent)).toEqual(['Name', 'Qty']);
		expect(box('Name').checked && box('Qty').checked).toBe(true);
		expect(store.selectionOf().ranges).toEqual([parseRange('A2:B5')]);

		fireEvent.click(box('My data has headers'));
		expect([...list.querySelectorAll('label')].map((l) => l.textContent)).toEqual([
			'Column A',
			'Column B',
		]);
		expect(store.selectionOf().ranges).toEqual([parseRange('A1:B5')]);
	});

	it('removes rows duplicated in the checked columns and reports the result', () => {
		const wb = people();
		const { store, onClose } = open(wb);
		fireEvent.click(box('Qty'));
		fireEvent.click(okButton());
		expect(onClose).toHaveBeenCalled();
		expect(column(wb, 'A1:A5')).toEqual(['Name', 'Ann', 'Bob', null, null]);
		expect(column(wb, 'B1:B5')).toEqual(['Qty', 1, 2, null, null]);
		expect(column(wb, 'D1:D1')).toEqual(['apart']);
		expect(store.state.message).toBe(
			'2 duplicate values found and removed; 2 unique values remain.',
		);
		expect(store.selectionOf().ranges).toEqual([parseRange('A1:B3')]);
		expect(store.getSnapshot().history.undo).toBe('Remove Duplicates');
		store.undo();
		expect(column(wb, 'A1:A5')).toEqual(['Name', 'Ann', 'Bob', 'ann', 'Bob']);
	});

	it('uses singular wording and reports when nothing is found', () => {
		const wb = people();
		const first = open(wb);
		fireEvent.keyDown(box('Name'), { key: 'Enter' });
		expect(first.store.state.message).toBe(
			'1 duplicate value found and removed; 3 unique values remain.',
		);
		first.unmount();
		const second = open(wb);
		fireEvent.click(okButton());
		expect(second.store.state.message).toBe('No duplicate values found.');
		expect(second.store.isDirty).toBe(false);
	});

	it('Select All / Unselect All drive the OK button', () => {
		open(people());
		fireEvent.click(screen.getByRole('button', { name: 'Unselect All' }));
		expect(box('Name').checked || box('Qty').checked).toBe(false);
		expect(okButton().disabled).toBe(true);
		fireEvent.click(screen.getByRole('button', { name: 'Select All' }));
		expect(box('Name').checked && box('Qty').checked).toBe(true);
		expect(okButton().disabled).toBe(false);
	});

	it('cancels with Escape and is disabled when read-only', () => {
		const wb = people();
		const { store, onClose, unmount } = open(wb);
		fireEvent.keyDown(box('Name'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
		expect(column(wb, 'A1:A5')).toEqual(['Name', 'Ann', 'Bob', 'ann', 'Bob']);
		unmount();
		open(wb, 'A2', true);
		expect(okButton().disabled).toBe(true);
	});

	it('reports merged cells without changing anything', () => {
		const wb = people();
		wb.sheets[0].merges.push(parseRange('A4:B4')!);
		const { store } = open(wb, 'A1:B5');
		fireEvent.click(okButton());
		expect(store.state.message).toBe("We can't do that to a merged cell.");
		expect(store.isDirty).toBe(false);
	});
});

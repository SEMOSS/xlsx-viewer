import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, getCell } from 'xlsx-viewer-core';

import { renderDialog, typeCells } from './__tests__/render-dialog';
import { NameManagerDialog } from './NameManagerDialog';

function book() {
	const wb = createWorkbook({ sheetNames: ['Data', 'Calc'] });
	typeCells(wb, { A1: 'Apple', B1: '5', B2: '7' });
	typeCells(wb, { A1: '=Rate*2' }, 1);
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Data!$B$1', comment: 'Base rate' },
		{ name: 'Items', formula: 'Data!$A$1:$B$1' },
		{ name: '_xlnm.Print_Area', formula: 'Data!$A$1:$B$2', localSheet: 0, builtin: true },
		{ name: 'Secret', formula: '1', hidden: true },
	);
	return wb;
}

const grid = () => screen.getByRole('grid', { name: 'Names' });
const bodyRows = () => within(grid()).getAllByRole('row').slice(1);
const rowNamed = (name: string) => bodyRows().find((r) => r.textContent?.startsWith(name))!;
const refersTo = () => screen.getByLabelText('Refers to:') as HTMLInputElement;
const button = (name: string) => screen.getByRole('button', { name }) as HTMLButtonElement;
const closeButton = () =>
	screen.getAllByRole('button', { name: 'Close' }).find((b) => b.hasAttribute('data-primary'))!;
const selectedName = () =>
	bodyRows()
		.find((r) => r.getAttribute('aria-selected') === 'true')
		?.querySelector('[role="gridcell"]')?.textContent;

describe('NameManagerDialog', () => {
	it('lists visible names sorted, with value, reference, scope and comment', () => {
		renderDialog(NameManagerDialog, { workbook: book() });
		expect(screen.getByRole('dialog', { name: 'Name Manager' })).toBeTruthy();
		const headers = within(grid())
			.getAllByRole('columnheader')
			.map((h) => h.textContent);
		expect(headers).toEqual(['Name', 'Value', 'Refers To', 'Scope', 'Comment']);
		const cells = bodyRows().map((r) =>
			within(r)
				.getAllByRole('gridcell')
				.map((c) => c.textContent),
		);
		expect(cells).toEqual([
			['Items', '{"Apple","5"}', '=Data!$A$1:$B$1', 'Workbook', ''],
			['Print_Area', '{"Apple","5";"","7"}', '=Data!$A$1:$B$2', 'Data', ''],
			['Rate', '5', '=Data!$B$1', 'Workbook', 'Base rate'],
		]);
		expect(selectedName()).toBe('Items');
		expect(refersTo().value).toBe('=Data!$A$1:$B$1');
		expect(document.activeElement).toBe(grid());
	});

	it('selects with the keyboard and the mouse, and edits Refers to in place', () => {
		const { store } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.keyDown(grid(), { key: 'ArrowDown' });
		expect(selectedName()).toBe('Print_Area');
		fireEvent.keyDown(grid(), { key: 'End' });
		expect(selectedName()).toBe('Rate');
		fireEvent.click(rowNamed('Items'));
		expect(refersTo().value).toBe('=Data!$A$1:$B$1');

		expect(button('Commit edit').disabled).toBe(true);
		fireEvent.change(refersTo(), { target: { value: '=Data!$A$1' } });
		fireEvent.click(button('Cancel edit'));
		expect(refersTo().value).toBe('=Data!$A$1:$B$1');

		fireEvent.change(refersTo(), { target: { value: '=$A$1:$A$2' } });
		fireEvent.click(button('Commit edit'));
		expect(store.workbook.definedNames.find((n) => n.name === 'Items')?.formula).toBe(
			'Data!$A$1:$A$2',
		);
		expect(store.getSnapshot().history.undo).toBe('Edit Name');
		expect(rowNamed('Items').textContent).toContain('{"Apple";""}');

		fireEvent.change(refersTo(), { target: { value: '=SUM(' } });
		fireEvent.keyDown(refersTo(), { key: 'Enter' });
		expect(screen.getByRole('alert').textContent).toBe("There's a problem with this formula.");
		fireEvent.change(refersTo(), { target: { value: '=Data!$B$2' } });
		fireEvent.keyDown(refersTo(), { key: 'Enter' });
		expect(store.workbook.definedNames.find((n) => n.name === 'Items')?.formula).toBe('Data!$B$2');
	});

	it('deletes a name after confirmation, as an undoable step', () => {
		const { store } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.click(rowNamed('Print_Area'));
		fireEvent.click(button('Delete'));
		const ask = screen.getByRole('alertdialog');
		expect(ask.textContent).toContain('Are you sure you want to delete the name Print_Area?');
		fireEvent.click(within(ask).getByRole('button', { name: 'Cancel' }));
		expect(screen.queryByRole('alertdialog')).toBeNull();
		expect(store.workbook.definedNames).toHaveLength(4);

		fireEvent.keyDown(grid(), { key: 'Delete' });
		fireEvent.click(within(screen.getByRole('alertdialog')).getByRole('button', { name: 'OK' }));
		expect(store.workbook.definedNames.some((n) => n.builtin)).toBe(false);
		expect(store.getSnapshot().history.undo).toBe('Delete Name');
		expect(bodyRows().map((r) => r.querySelector('[role="gridcell"]')?.textContent)).toEqual([
			'Items',
			'Rate',
		]);
		expect(selectedName()).toBe('Rate');
		expect(document.activeElement).toBe(grid());
	});

	it('opens New Name on top and selects the added name', () => {
		const { store, onClose } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.click(button('New...'));
		const editor = screen.getByRole('dialog', { name: 'New Name' });
		fireEvent.keyDown(within(editor).getByLabelText('Name:'), { key: 'Escape' });
		expect(screen.queryByRole('dialog', { name: 'New Name' })).toBeNull();
		expect(onClose).not.toHaveBeenCalled();

		fireEvent.click(button('New...'));
		fireEvent.change(screen.getByLabelText('Name:'), { target: { value: 'Tax' } });
		const nested = screen.getByRole('dialog', { name: 'New Name' });
		fireEvent.change(within(nested).getByLabelText('Refers to:'), {
			target: { value: '=Data!$B$2' },
		});
		fireEvent.click(button('OK'));
		expect(screen.queryByRole('dialog', { name: 'New Name' })).toBeNull();
		expect(store.workbook.definedNames).toContainEqual({ name: 'Tax', formula: 'Data!$B$2' });
		expect(store.getSnapshot().history.undo).toBe('Define Name');
		expect(selectedName()).toBe('Tax');
		expect(rowNamed('Tax').textContent).toContain('7');
	});

	it('edits a name in the editor and renames its uses', () => {
		const { store } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.doubleClick(rowNamed('Rate'));
		const editor = screen.getByRole('dialog', { name: 'Edit Name' });
		const name = within(editor).getByLabelText('Name:') as HTMLInputElement;
		expect(name.value).toBe('Rate');
		expect(within(editor).getByLabelText('Comment:')).toHaveProperty('value', 'Base rate');
		expect(
			(within(editor).getByRole('combobox', { name: 'Scope' }) as HTMLSelectElement).disabled,
		).toBe(true);
		fireEvent.change(name, { target: { value: 'Factor' } });
		fireEvent.keyDown(name, { key: 'Enter' });
		expect(screen.queryByRole('dialog', { name: 'Edit Name' })).toBeNull();
		expect(getCell(store.workbook.sheets[1], 0, 0)?.formula).toBe('Factor*2');
		expect(selectedName()).toBe('Factor');
		expect(store.getSnapshot().history.undo).toBe('Edit Name');
		store.undo();
		expect(getCell(store.workbook.sheets[1], 0, 0)?.formula).toBe('Rate*2');
	});

	it('asks to keep an edited reference before closing', () => {
		const { store, onClose } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.change(refersTo(), { target: { value: '=Data!$A$1' } });
		fireEvent.click(closeButton());
		const ask = screen.getByRole('alertdialog');
		expect(ask.textContent).toContain(
			'Do you want to save the changes you made to the name reference?',
		);
		fireEvent.click(within(ask).getByRole('button', { name: 'No' }));
		expect(onClose).toHaveBeenCalledTimes(1);
		expect(store.workbook.definedNames.find((n) => n.name === 'Items')?.formula).toBe(
			'Data!$A$1:$B$1',
		);
	});

	it('saves the edited reference when the answer is Yes, and Enter closes', () => {
		const { store, onClose } = renderDialog(NameManagerDialog, { workbook: book() });
		fireEvent.change(refersTo(), { target: { value: '=Data!$A$1' } });
		fireEvent.keyDown(grid(), { key: 'Escape' });
		fireEvent.click(within(screen.getByRole('alertdialog')).getByRole('button', { name: 'Yes' }));
		expect(onClose).toHaveBeenCalledTimes(1);
		expect(store.workbook.definedNames.find((n) => n.name === 'Items')?.formula).toBe('Data!$A$1');

		onClose.mockClear();
		fireEvent.keyDown(grid(), { key: 'Enter' });
		expect(onClose).toHaveBeenCalledTimes(1);
	});

	it('only lists names when read-only', () => {
		const { store } = renderDialog(NameManagerDialog, { workbook: book(), readOnly: true });
		expect(button('New...').disabled).toBe(true);
		expect(button('Edit...').disabled).toBe(true);
		expect(button('Delete').disabled).toBe(true);
		expect(refersTo().readOnly).toBe(true);
		fireEvent.keyDown(grid(), { key: 'Delete' });
		fireEvent.doubleClick(rowNamed('Rate'));
		expect(screen.queryByRole('alertdialog')).toBeNull();
		expect(screen.queryByRole('dialog', { name: 'Edit Name' })).toBeNull();
		expect(store.workbook.definedNames).toHaveLength(4);
	});

	it('shows an empty list for a workbook without names', () => {
		renderDialog(NameManagerDialog);
		expect(bodyRows()).toEqual([]);
		expect(button('Edit...').disabled).toBe(true);
		expect(refersTo().disabled).toBe(true);
	});
});

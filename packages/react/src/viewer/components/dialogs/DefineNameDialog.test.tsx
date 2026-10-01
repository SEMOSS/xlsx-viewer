import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { DefineNameDialog } from './DefineNameDialog';

function book() {
	const wb = createWorkbook({ sheetNames: ['Sales Data', 'Other'] });
	typeCells(wb, { B1: 'Unit Price', B2: '10', B3: '12' });
	wb.definedNames.push({ name: 'Existing', formula: "'Sales Data'!$A$1" });
	return wb;
}

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;

describe('DefineNameDialog', () => {
	it('prefills an empty name and the active cell for an unlabelled selection', () => {
		renderDialog(DefineNameDialog, { workbook: book() });
		expect(screen.getByRole('dialog', { name: 'New Name' })).toBeTruthy();
		expect(field('Name:').value).toBe('');
		expect(field('Refers to:').value).toBe("='Sales Data'!$A$1");
	});

	it('suggests the label above the selection and adds the name as one undoable step', () => {
		const store = new ViewerStore(book());
		selectRange(store, 'B2:B3');
		const { onClose } = renderDialog(DefineNameDialog, { store });
		expect(field('Name:').value).toBe('Unit_Price');
		expect((screen.getByRole('combobox', { name: 'Scope' }) as HTMLSelectElement).value).toBe(
			'workbook',
		);
		expect(field('Refers to:').value).toBe("='Sales Data'!$B$2:$B$3");
		fireEvent.change(field('Comment:'), { target: { value: 'List price' } });
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.definedNames).toContainEqual({
			name: 'Unit_Price',
			formula: "'Sales Data'!$B$2:$B$3",
			comment: 'List price',
		});
		expect(store.getSnapshot().history.undo).toBe('Define Name');
		store.undo();
		expect(store.workbook.definedNames.map((n) => n.name)).toEqual(['Existing']);
	});

	it('keeps the dialog open on errors', () => {
		const { store, onClose } = renderDialog(DefineNameDialog, { workbook: book() });
		fireEvent.change(field('Name:'), { target: { value: 'existing' } });
		fireEvent.keyDown(field('Name:'), { key: 'Enter' });
		expect(screen.getByRole('alert').textContent).toBe(
			'The name already exists. Enter a unique name.',
		);
		fireEvent.change(field('Name:'), { target: { value: 'B2' } });
		fireEvent.keyDown(field('Name:'), { key: 'Enter' });
		expect(screen.getByRole('alert').textContent).toContain(
			'The name that you entered is not valid.',
		);
		expect(onClose).not.toHaveBeenCalled();
		expect(store.workbook.definedNames).toHaveLength(1);
		expect(store.isDirty).toBe(false);
	});

	it('creates sheet-scoped names', () => {
		const { store } = renderDialog(DefineNameDialog, { workbook: book() });
		fireEvent.change(field('Name:'), { target: { value: 'Existing' } });
		fireEvent.change(screen.getByRole('combobox', { name: 'Scope' }), { target: { value: '1' } });
		fireEvent.change(field('Refers to:'), { target: { value: '=Other!$C$1' } });
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(store.workbook.definedNames).toContainEqual({
			name: 'Existing',
			formula: 'Other!$C$1',
			localSheet: 1,
		});
	});

	it('cancels with Escape and disables OK when read-only', () => {
		const { store, onClose } = renderDialog(DefineNameDialog, { workbook: book(), readOnly: true });
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
		fireEvent.keyDown(field('Name:'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.definedNames).toHaveLength(1);
	});
});

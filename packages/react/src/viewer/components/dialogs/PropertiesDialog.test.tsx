import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, typeCells } from './__tests__/render-dialog';
import { PropertiesDialog } from './PropertiesDialog';

const CREATED = '2024-01-01T10:00:00Z';

function storeWithProperties(readOnly = false) {
	const wb = createWorkbook({ sheetNames: ['Data', 'Calc', 'Notes'] });
	typeCells(wb, { A1: 'Item', A2: '5', A3: '=A2*2' });
	typeCells(wb, { B2: 'x' }, 1);
	wb.properties = {
		title: 'Budget',
		creator: 'Ann',
		company: 'Contoso',
		created: CREATED,
		lastModifiedBy: 'Bob',
	};
	wb.dirty = false;
	return new ViewerStore(wb, { readOnly });
}

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;
const change = (label: string, value: string) =>
	fireEvent.change(field(label), { target: { value } });

describe('PropertiesDialog', () => {
	it('edits the summary properties outside the undo history', () => {
		const store = storeWithProperties();
		const { onClose } = renderDialog(PropertiesDialog, { store });
		expect(field('Title:').value).toBe('Budget');
		expect(field('Author:').value).toBe('Ann');
		expect(field('Company:').value).toBe('Contoso');
		expect(field('Manager:').value).toBe('');
		change('Title:', '  ');
		change('Subject:', ' Q3 numbers ');
		change('Keywords:', 'budget; 2024');
		change('Comments:', 'Line 1\nLine 2');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		const props = store.workbook.properties;
		expect(props.title).toBeUndefined();
		expect(props.subject).toBe('Q3 numbers');
		expect(props.keywords).toBe('budget; 2024');
		expect(props.description).toBe('Line 1\nLine 2');
		expect(props.creator).toBe('Ann');
		expect(props.created).toBe(CREATED);
		expect(store.workbook.dirty).toBe(true);
		expect(store.isDirty).toBe(true);
		expect(store.getSnapshot().history.undo).toBeNull();
	});

	it('shows statistics on their own tab', () => {
		renderDialog(PropertiesDialog, { store: storeWithProperties() });
		fireEvent.click(screen.getByRole('tab', { name: 'Statistics' }));
		const value = (label: string) => screen.getByText(label).nextElementSibling?.textContent;
		expect(value('Created:')).toBe(new Date(CREATED).toLocaleString('en'));
		expect(value('Modified:')).toBe('Not known');
		expect(value('Last modified by:')).toBe('Bob');
		expect(value('Sheets:')).toBe('3');
		expect(value('Cells with values:')).toBe('4');
		expect(value('Formulas:')).toBe('1');
		expect(screen.queryByLabelText('Title:')).toBeNull();
	});

	it('runs OK on Enter in a field but not in Comments', () => {
		const store = storeWithProperties();
		const { onClose } = renderDialog(PropertiesDialog, { store });
		change('Category:', 'Finance');
		fireEvent.keyDown(field('Comments:'), { key: 'Enter' });
		expect(onClose).not.toHaveBeenCalled();
		fireEvent.keyDown(field('Category:'), { key: 'Enter' });
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.properties.category).toBe('Finance');
	});

	it('does not mark the workbook changed when nothing changed', () => {
		const store = storeWithProperties();
		const { onClose } = renderDialog(PropertiesDialog, { store });
		change('Title:', 'Budget ');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
		expect(store.workbook.dirty).toBe(false);
	});

	it('cancels with Escape without changing anything', () => {
		const store = storeWithProperties();
		const { onClose } = renderDialog(PropertiesDialog, { store });
		change('Title:', 'Other');
		fireEvent.keyDown(field('Title:'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.properties.title).toBe('Budget');
		expect(store.isDirty).toBe(false);
	});

	it('shows read-only fields on a read-only workbook and OK only closes', () => {
		const store = storeWithProperties(true);
		const { onClose } = renderDialog(PropertiesDialog, { store });
		expect(field('Title:').readOnly).toBe(true);
		expect(field('Comments:').readOnly).toBe(true);
		change('Title:', 'Changed anyway');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.properties.title).toBe('Budget');
		expect(store.isDirty).toBe(false);
	});
});

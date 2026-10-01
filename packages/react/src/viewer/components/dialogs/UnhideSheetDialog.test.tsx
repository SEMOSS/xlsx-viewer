import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { renderDialog } from './__tests__/render-dialog';
import { UnhideSheetDialog } from './UnhideSheetDialog';

function workbookWithHidden() {
	const wb = createWorkbook({ sheetNames: ['Data', 'Calc', 'Secret', 'Archive'] });
	wb.sheets[2].state = 'hidden';
	wb.sheets[3].state = 'hidden';
	wb.sheets[1].state = 'veryHidden';
	return wb;
}

describe('UnhideSheetDialog', () => {
	it('lists only hidden (not very hidden) sheets and unhides the chosen one', () => {
		const { store, onClose } = renderDialog(UnhideSheetDialog, { workbook: workbookWithHidden() });
		const options = screen.getAllByRole('option');
		expect(options.map((o) => o.textContent)).toEqual(['Secret', 'Archive']);
		fireEvent.click(options[1]);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.sheets[3].state).toBe('visible');
		expect(store.workbook.sheets[2].state).toBe('hidden');
		expect(store.state.activeSheet).toBe(3);
	});

	it('runs OK on Enter in the list and supports keyboard selection', () => {
		const { store } = renderDialog(UnhideSheetDialog, { workbook: workbookWithHidden() });
		const list = screen.getByRole('listbox', { name: 'Unhide sheet' });
		fireEvent.keyDown(list, { key: 'ArrowDown' });
		expect(screen.getAllByRole('option')[1].getAttribute('aria-selected')).toBe('true');
		fireEvent.keyDown(list, { key: 'Enter' });
		expect(store.workbook.sheets[3].state).toBe('visible');
	});

	it('cancels with Escape without changing anything', () => {
		const { store, onClose } = renderDialog(UnhideSheetDialog, { workbook: workbookWithHidden() });
		fireEvent.keyDown(screen.getByRole('listbox'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.workbook.sheets[2].state).toBe('hidden');
		expect(store.isDirty).toBe(false);
	});

	it('disables OK when nothing is hidden', () => {
		renderDialog(UnhideSheetDialog);
		expect(screen.getByText('There are no hidden sheets.')).toBeTruthy();
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
	});
});

import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, getCell, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { DeleteCellsDialog, InsertCellsDialog } from './InsertDeleteCellsDialog';

function seeded() {
	const wb = createWorkbook();
	typeCells(wb, { A2: 'a2', B2: 'b2', C2: 'c2', B3: 'b3', D1: '=B3' });
	return wb;
}

const value = (wb: ReturnType<typeof createWorkbook>, ref: string) => {
	const r = parseRange(ref)!;
	return getCell(wb.sheets[0], r.top, r.left)?.value ?? null;
};

const radio = (name: string) => screen.getByRole('radio', { name }) as HTMLInputElement;

describe('InsertCellsDialog', () => {
	it('offers Excel choices with shift down by default and shifts cells right on request', () => {
		const wb = seeded();
		const { store, onClose } = renderDialog(InsertCellsDialog, { workbook: wb });
		expect(screen.getByRole('dialog', { name: 'Insert' })).toBeTruthy();
		expect(screen.getByRole('radiogroup', { name: 'Insert' })).toBeTruthy();
		expect(screen.getAllByRole('radio').map((r) => r.parentElement?.textContent)).toEqual([
			'Shift cells right',
			'Shift cells down',
			'Entire row',
			'Entire column',
		]);
		expect(radio('Shift cells down').checked).toBe(true);
		selectRange(store, 'B2');
		fireEvent.click(radio('Shift cells right'));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(['A2', 'B2', 'C2', 'D2'].map((a) => value(wb, a))).toEqual(['a2', null, 'b2', 'c2']);
		expect(value(wb, 'B3')).toBe('b3');
	});

	it('defaults to shift right for tall selections and runs OK on Enter', () => {
		const wb = seeded();
		const store = new ViewerStore(wb);
		selectRange(store, 'B2:B3');
		const { onClose } = renderDialog(InsertCellsDialog, { store });
		expect(radio('Shift cells right').checked).toBe(true);
		expect(document.activeElement).toBe(radio('Shift cells right'));
		fireEvent.keyDown(radio('Shift cells right'), { key: 'Enter' });
		expect(onClose).toHaveBeenCalled();
		expect(['B2', 'C2', 'C3'].map((a) => value(wb, a))).toEqual([null, 'b2', 'b3']);
	});

	it('suggests entire columns or rows for whole-axis selections', () => {
		const store = new ViewerStore(seeded());
		selectRange(store, 'B:C');
		const first = renderDialog(InsertCellsDialog, { store });
		expect(radio('Entire column').checked).toBe(true);
		first.unmount();
		selectRange(store, '2:3');
		renderDialog(DeleteCellsDialog, { store });
		expect(radio('Entire row').checked).toBe(true);
	});

	it('inserts entire rows', () => {
		const wb = seeded();
		const { store } = renderDialog(InsertCellsDialog, { workbook: wb });
		selectRange(store, 'B2');
		fireEvent.click(radio('Entire row'));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(['A2', 'A3', 'C3', 'B4'].map((a) => value(wb, a))).toEqual([null, 'a2', 'c2', 'b3']);
	});

	it('cancels with Escape without changing anything', () => {
		const wb = seeded();
		const { store, onClose } = renderDialog(InsertCellsDialog, { workbook: wb });
		fireEvent.keyDown(radio('Shift cells down'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
		expect(value(wb, 'B2')).toBe('b2');
	});
});

describe('DeleteCellsDialog', () => {
	it('shifts cells up by default and adjusts references', () => {
		const wb = seeded();
		const { store } = renderDialog(DeleteCellsDialog, { workbook: wb });
		expect(screen.getByRole('dialog', { name: 'Delete' })).toBeTruthy();
		expect(screen.getAllByRole('radio').map((r) => r.parentElement?.textContent)).toEqual([
			'Shift cells left',
			'Shift cells up',
			'Entire row',
			'Entire column',
		]);
		selectRange(store, 'B2');
		expect(radio('Shift cells up').checked).toBe(true);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(value(wb, 'B2')).toBe('b3');
		expect(getCell(wb.sheets[0], 0, 3)?.formula).toBe('B2');
		expect(store.getSnapshot().history.undo).toBe('Delete Cells');
	});

	it('shows refusals in the status bar and changes nothing', () => {
		const wb = seeded();
		wb.sheets[0].merges.push(parseRange('A5:C5')!);
		const { store } = renderDialog(DeleteCellsDialog, { workbook: wb });
		selectRange(store, 'B2');
		fireEvent.click(radio('Shift cells up'));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(store.state.message).toBe("We can't do that to a merged cell.");
		expect(value(wb, 'B2')).toBe('b2');
		expect(store.isDirty).toBe(false);
	});

	it('removes entire columns', () => {
		const wb = seeded();
		const { store } = renderDialog(DeleteCellsDialog, { workbook: wb });
		selectRange(store, 'B2');
		fireEvent.click(radio('Entire column'));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(['A2', 'B2', 'B3'].map((a) => value(wb, a))).toEqual(['a2', 'c2', null]);
	});

	it('cancels with the Cancel button without changing anything', () => {
		const wb = seeded();
		const { store, onClose } = renderDialog(DeleteCellsDialog, { workbook: wb });
		fireEvent.click(radio('Shift cells left'));
		fireEvent.click(screen.getByRole('button', { name: 'Cancel' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.isDirty).toBe(false);
		expect(value(wb, 'B2')).toBe('b2');
	});

	it('disables OK in read-only mode', () => {
		renderDialog(DeleteCellsDialog, { workbook: seeded(), readOnly: true });
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
	});
});

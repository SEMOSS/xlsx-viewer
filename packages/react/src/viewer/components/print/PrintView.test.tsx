import { act, fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, putCell } from 'xlsx-viewer-core';

import { ShellContext, type ShellActions } from '../../shell-context';
import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { PrintView } from './PrintView';

/** A sheet of 100 rows by 15 columns: 3 pages tall, 2 wide on Letter portrait. */
function setup() {
	const wb = createWorkbook({ sheetNames: ['Data', 'Notes'] });
	putCell(wb.sheets[0], 0, 0, { value: 'Top' });
	putCell(wb.sheets[0], 99, 14, { value: 'Bottom' });
	putCell(wb.sheets[1], 0, 0, { value: 'Note' });
	const store = new ViewerStore(wb);
	const shell: ShellActions = {
		openDialog: vi.fn(),
		save: vi.fn(),
		download: vi.fn(),
		print: vi.fn(),
		notify: vi.fn(),
		themeMode: 'light',
	};
	render(
		<ViewerStoreContext.Provider value={store}>
			<ShellContext.Provider value={shell}>
				<PrintView fileName='Book.xlsx' />
			</ShellContext.Provider>
		</ViewerStoreContext.Provider>,
	);
	return { store, shell };
}

const currentPage = () =>
	(screen.getByRole('textbox', { name: 'Current Page' }) as HTMLInputElement).value;

describe('PrintView', () => {
	it('previews every page and moves between them', () => {
		setup();
		expect(screen.getByText('of 6')).toBeTruthy();
		expect(currentPage()).toBe('1');
		expect(screen.getByRole('img', { name: 'Page 1' })).toBeTruthy();
		expect(screen.getByText('Top')).toBeTruthy();
		fireEvent.click(screen.getByRole('button', { name: 'Next Page' }));
		expect(currentPage()).toBe('2');
		expect(screen.getByRole('img', { name: 'Page 2' })).toBeTruthy();
		expect(screen.queryByText('Top')).toBeNull();
		fireEvent.change(screen.getByRole('textbox', { name: 'Current Page' }), {
			target: { value: '6' },
		});
		expect(screen.getByRole('img', { name: 'Page 6' })).toBeTruthy();
		expect(screen.getByText('Bottom')).toBeTruthy();
		expect((screen.getByRole('button', { name: 'Next Page' }) as HTMLButtonElement).disabled).toBe(
			true,
		);
		fireEvent.click(screen.getByRole('button', { name: 'Previous Page' }));
		expect(currentPage()).toBe('5');
	});

	it('applies settings to the sheet as undoable edits', () => {
		const { store } = setup();
		fireEvent.change(screen.getByLabelText('Scaling'), { target: { value: 'fitSheet' } });
		expect(store.activeSheet.pageSetup).toMatchObject({
			fitToPage: true,
			fitToWidth: 1,
			fitToHeight: 1,
		});
		expect(screen.getByText('of 1')).toBeTruthy();
		fireEvent.change(screen.getByLabelText('Orientation'), { target: { value: 'landscape' } });
		expect(store.activeSheet.pageSetup?.orientation).toBe('landscape');
		fireEvent.change(screen.getByLabelText('Margins'), { target: { value: 'narrow' } });
		expect(store.activeSheet.pageMargins?.left).toBe(0.25);
		fireEvent.change(screen.getByLabelText('Paper size'), { target: { value: '9' } });
		expect(store.activeSheet.pageSetup?.paperSize).toBe(9);
		act(() => store.undo());
		act(() => store.undo());
		act(() => store.undo());
		expect(store.activeSheet.pageSetup?.orientation).toBeUndefined();
		expect(screen.getByText('of 1')).toBeTruthy();
		act(() => store.undo());
		expect(screen.getByText('of 6')).toBeTruthy();
	});

	it('prints the entire workbook or the selection', () => {
		const { store } = setup();
		fireEvent.change(screen.getByLabelText('Print what'), { target: { value: 'workbook' } });
		expect(screen.getByText('of 7')).toBeTruthy();
		act(() =>
			store.setSelection({
				active: { row: 0, col: 0 },
				anchor: { row: 0, col: 0 },
				ranges: [{ top: 0, left: 0, bottom: 1, right: 1 }],
			}),
		);
		fireEvent.change(screen.getByLabelText('Print what'), { target: { value: 'selection' } });
		expect(screen.getByText('of 1')).toBeTruthy();
	});

	it('opens Page Setup for custom settings', () => {
		const { shell } = setup();
		fireEvent.change(screen.getByLabelText('Margins'), { target: { value: 'dialog' } });
		expect(shell.openDialog).toHaveBeenCalledWith('pageSetup', { tab: 'margins' });
		fireEvent.click(screen.getByRole('button', { name: 'Page Setup' }));
		expect(shell.openDialog).toHaveBeenCalledWith('pageSetup');
	});
});

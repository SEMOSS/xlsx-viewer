import { act, fireEvent, render, screen } from '@testing-library/react';
import {
	createTable,
	createWorkbook,
	EditTransaction,
	parseRange,
	setCellInput,
} from 'xlsx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ShellContext, type ShellActions } from '../../../shell-context';
import { ViewerStoreContext } from '../../../state/use-store';
import { ViewerStore } from '../../../state/viewer-store';
import { Ribbon } from '../Ribbon';

function setup() {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	[
		['Item', 'Qty'],
		['Pen', '3'],
		['Ink', '5'],
	].forEach((r, i) => r.forEach((v, j) => setCellInput(tx, wb, 0, i, j, v)));
	createTable(tx, wb, 0, parseRange('A1:B3')!);
	tx.commit();
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
				<Ribbon onOpenBackstage={() => {}} />
			</ShellContext.Provider>
		</ViewerStoreContext.Provider>,
	);
	return store;
}

const select = (store: ViewerStore, row: number, col: number) =>
	act(() =>
		store.setSelection({
			active: { row, col },
			anchor: { row, col },
			ranges: [{ top: row, left: col, bottom: row, right: col }],
		}),
	);

describe('Table Design tab', () => {
	it('appears only while the active cell is in a table', () => {
		const store = setup();
		select(store, 6, 4);
		expect(screen.queryByRole('tab', { name: 'Table Design' })).toBeNull();
		select(store, 1, 1);
		expect(screen.getByRole('tab', { name: 'Table Design' })).toBeTruthy();
	});

	it('toggles the total row and renames the table', () => {
		const store = setup();
		select(store, 1, 0);
		fireEvent.click(screen.getByRole('tab', { name: 'Table Design' }));
		fireEvent.click(screen.getByRole('checkbox', { name: 'Total Row' }));
		expect(store.activeSheet.tables[0].totalsRowCount).toBe(1);
		const name = screen.getByRole('textbox', { name: 'Table Name' });
		fireEvent.change(name, { target: { value: 'Stock' } });
		fireEvent.keyDown(name, { key: 'Enter' });
		expect(store.activeSheet.tables[0].name).toBe('Stock');
	});

	it('returns to Home when the selection leaves the table', () => {
		const store = setup();
		select(store, 1, 0);
		fireEvent.click(screen.getByRole('tab', { name: 'Table Design' }));
		select(store, 8, 8);
		expect(screen.queryByRole('tab', { name: 'Table Design' })).toBeNull();
		expect(store.state.ribbonTab).toBe('home');
	});
});

import { act, fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, getPrintArea, putCell } from 'xlsx-viewer-core';

import { ShellContext, type ShellActions } from '../../../shell-context';
import { ViewerStoreContext } from '../../../state/use-store';
import { ViewerStore } from '../../../state/viewer-store';
import { Ribbon } from '../Ribbon';

function setup() {
	const wb = createWorkbook();
	putCell(wb.sheets[0], 0, 0, { value: 'x' });
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
	fireEvent.click(screen.getByRole('tab', { name: 'Page Layout' }));
	return { store, shell };
}

const menu = (button: string, item: RegExp | string) => {
	fireEvent.click(screen.getByRole('button', { name: button }));
	const items = [
		...screen.queryAllByRole('menuitem'),
		...screen.queryAllByRole('menuitemcheckbox'),
	];
	const hit = items.find((el) =>
		typeof item === 'string' ? el.textContent === item : item.test(el.textContent ?? ''),
	);
	if (!hit) {
		throw new Error(`no menu item ${String(item)}`);
	}
	fireEvent.click(hit);
};

const select = (store: ViewerStore, row: number, col: number) =>
	act(() =>
		store.setSelection({
			active: { row, col },
			anchor: { row, col },
			ranges: [{ top: row, left: col, bottom: row + 1, right: col + 1 }],
		}),
	);

describe('Page Layout tab', () => {
	it('sits between Insert and Formulas', () => {
		setup();
		const tabs = screen.getAllByRole('tab').map((t) => t.textContent);
		expect(tabs.indexOf('Page Layout')).toBe(tabs.indexOf('Insert') + 1);
		expect(tabs.indexOf('Formulas')).toBe(tabs.indexOf('Page Layout') + 1);
	});

	it('changes orientation and margins, undoably', () => {
		const { store } = setup();
		menu('Orientation', 'Landscape Orientation');
		expect(store.activeSheet.pageSetup?.orientation).toBe('landscape');
		menu('Margins', /^Wide/);
		expect(store.activeSheet.pageMargins).toEqual({
			top: 1,
			bottom: 1,
			left: 1,
			right: 1,
			header: 0.5,
			footer: 0.5,
		});
		act(() => store.undo());
		expect(store.activeSheet.pageMargins).toBeUndefined();
		act(() => store.undo());
		expect(store.activeSheet.pageSetup?.orientation).toBeUndefined();
		act(() => store.redo());
		expect(store.activeSheet.pageSetup?.orientation).toBe('landscape');
	});

	it('sets the paper size, print area and page breaks', () => {
		const { store, shell } = setup();
		menu('Size', /^A4/);
		expect(store.activeSheet.pageSetup?.paperSize).toBe(9);
		select(store, 2, 1);
		menu('Print Area', 'Set Print Area');
		expect(getPrintArea(store.workbook, 0)).toEqual([{ top: 2, left: 1, bottom: 3, right: 2 }]);
		menu('Print Area', 'Clear Print Area');
		expect(getPrintArea(store.workbook, 0)).toEqual([]);
		menu('Breaks', 'Insert Page Break');
		expect(store.activeSheet.rowBreaks).toEqual([2]);
		expect(store.activeSheet.colBreaks).toEqual([1]);
		menu('Breaks', 'Remove Page Break');
		expect(store.activeSheet.rowBreaks).toEqual([]);
		fireEvent.click(screen.getByRole('button', { name: 'Print Titles' }));
		expect(shell.openDialog).toHaveBeenCalledWith('pageSetup', { tab: 'sheet' });
		menu('Margins', 'Custom Margins...');
		expect(shell.openDialog).toHaveBeenCalledWith('pageSetup', { tab: 'margins' });
	});

	it('scales to fit and toggles sheet options', () => {
		const { store } = setup();
		const scale = screen.getByRole('textbox', { name: 'Scale' }) as HTMLInputElement;
		expect(scale.disabled).toBe(false);
		fireEvent.change(scale, { target: { value: '75' } });
		fireEvent.keyDown(scale, { key: 'Enter' });
		expect(store.activeSheet.pageSetup).toMatchObject({ scale: 75, fitToPage: false });
		fireEvent.change(screen.getByRole('combobox', { name: 'Width' }), { target: { value: '1' } });
		expect(store.activeSheet.pageSetup).toMatchObject({
			fitToPage: true,
			fitToWidth: 1,
			fitToHeight: 0,
		});
		expect((screen.getByRole('textbox', { name: 'Scale' }) as HTMLInputElement).disabled).toBe(
			true,
		);
		fireEvent.change(screen.getByRole('combobox', { name: 'Width' }), { target: { value: '0' } });
		expect(store.activeSheet.pageSetup?.fitToPage).toBe(false);
		fireEvent.click(screen.getByRole('checkbox', { name: 'Gridlines Print' }));
		expect(store.activeSheet.printOptions?.gridLines).toBe(true);
		fireEvent.click(screen.getByRole('checkbox', { name: 'Headings Print' }));
		expect(store.activeSheet.printOptions?.headings).toBe(true);
		fireEvent.click(screen.getByRole('checkbox', { name: 'Gridlines View' }));
		expect(store.activeSheet.view.showGridLines).toBe(false);
		act(() => store.undo());
		expect(store.activeSheet.printOptions).toEqual({ gridLines: true });
	});
});

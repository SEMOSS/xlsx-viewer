import { act, fireEvent, render, screen } from '@testing-library/react';
import { createWorkbook, MAX_COLS } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { stubLayout } from '../chrome/__tests__/layout-stubs';
import { Grid } from './Grid';

stubLayout();

function renderGrid() {
	const store = new ViewerStore(createWorkbook({ sheetNames: ['Budget'] }));
	render(
		<ViewerStoreContext.Provider value={store}>
			<Grid />
		</ViewerStoreContext.Provider>,
	);
	return store;
}

describe('Grid', () => {
	it('names the grid after the sheet and draws the row and column headers', () => {
		renderGrid();
		const grid = screen.getByRole('grid', { name: 'Sheet Budget' });
		expect(grid.querySelector('[data-col="0"]')?.textContent).toBe('A');
		expect(grid.querySelector('[data-row="0"]')?.textContent).toBe('1');
	});

	it('keeps the keyboard in the grid when a cell is clicked', () => {
		const store = renderGrid();
		const grid = screen.getByRole('grid');
		const sink = screen.getByLabelText('Spreadsheet');
		// jsdom has no layout: give the scroller a size and pointer capture.
		Object.defineProperty(grid, 'clientWidth', { value: 800 });
		Object.defineProperty(grid, 'clientHeight', { value: 400 });
		grid.setPointerCapture = () => undefined;
		sink.blur();
		const press = (button: number) => {
			const e = new MouseEvent('pointerdown', {
				bubbles: true,
				cancelable: true,
				button,
				clientX: 200,
				clientY: 70,
			});
			act(() => {
				grid.dispatchEvent(e);
			});
			return e;
		};
		// Cancelling the press stops the browser from moving focus off the sink.
		expect(press(0).defaultPrevented).toBe(true);
		expect(document.activeElement).toBe(sink);
		expect(store.selectionOf().active).not.toEqual({ row: 0, col: 0 });
		sink.blur();
		expect(press(2).defaultPrevented).toBe(true);
		expect(document.activeElement).toBe(sink);
	});

	it('selecting whole rows does not scroll sideways', async () => {
		const store = renderGrid();
		const grid = screen.getByRole('grid');
		Object.defineProperty(grid, 'clientWidth', { value: 800 });
		Object.defineProperty(grid, 'clientHeight', { value: 400 });
		// jsdom does not scroll: record what the grid asks for instead.
		const scrolledTo: number[] = [];
		Object.defineProperty(grid, 'scrollLeft', {
			get: () => 0,
			set: (v: number) => scrolledTo.push(v),
		});
		act(() => {
			store.setSelection({
				active: { row: 4, col: 0 },
				anchor: { row: 4, col: 0 },
				ranges: [{ top: 4, left: 0, bottom: 4, right: MAX_COLS - 1 }],
			});
		});
		await act(() => new Promise((r) => requestAnimationFrame(() => r(undefined))));
		expect(scrolledTo.every((left) => left === 0)).toBe(true);
	});

	it('draws overlays only in panes that show cells (one copy without frozen panes)', () => {
		renderGrid();
		const grid = screen.getByRole('grid');
		expect(grid.querySelectorAll('.xv-selection')).toHaveLength(1);
		expect(grid.querySelector('.xv-pane-body .xv-selection')).toBeTruthy();
	});

	it('starts an edit when a character is typed into the focus sink', () => {
		const store = renderGrid();
		const sink = screen.getByLabelText('Spreadsheet') as HTMLTextAreaElement;
		sink.value = '7';
		fireEvent.input(sink);
		expect(store.state.edit?.text).toBe('7');
		expect(sink.value).toBe('');
	});
});

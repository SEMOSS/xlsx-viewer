import { act, fireEvent, render, screen } from '@testing-library/react';
import { cellKey, createChartFromRange, createWorkbook, type ChartObject } from 'xlsx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ShellContext, type ShellActions } from '../../../shell-context';
import { ViewerStoreContext } from '../../../state/use-store';
import { ViewerStore } from '../../../state/viewer-store';
import { Ribbon } from '../Ribbon';

function setup() {
	const wb = createWorkbook();
	const rows: Array<Array<string | number | null>> = [
		[null, 'Revenue', 'Cost'],
		['Q1', 10, 4],
		['Q2', 20, 8],
		['Q3', 30, 12],
		['Q4', 40, 16],
	];
	rows.forEach((row, r) =>
		row.forEach((value, c) => {
			if (value !== null) {
				wb.sheets[0].cells.set(cellKey(r, c), { value });
			}
		}),
	);
	const chart: ChartObject = {
		kind: 'chart',
		id: 'chart-1',
		chart: createChartFromRange(wb, 0, { top: 0, left: 0, bottom: 4, right: 2 }, 'column'),
		anchor: {
			type: 'twoCell',
			from: { row: 0, col: 4, rowOff: 0, colOff: 0 },
			to: { row: 15, col: 11, rowOff: 0, colOff: 0 },
		},
	};
	wb.sheets[0].drawings = [chart];
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
	const model = () => (store.activeSheet.drawings[0] as ChartObject).chart;
	return { store, model };
}

describe('Chart Design tab', () => {
	it('appears while a chart is selected and returns to Home afterwards', () => {
		const { store } = setup();
		expect(screen.queryByRole('tab', { name: 'Chart Design' })).toBeNull();
		act(() => store.setUi({ selectedObject: 'chart-1' }));
		fireEvent.click(screen.getByRole('tab', { name: 'Chart Design' }));
		expect(store.state.ribbonTab).toBe('chartDesign');
		expect(screen.getByRole('button', { name: 'Switch Row/Column' })).toBeTruthy();
		act(() => store.setUi({ selectedObject: null }));
		expect(screen.queryByRole('tab', { name: 'Chart Design' })).toBeNull();
		expect(store.state.ribbonTab).toBe('home');
	});

	it('changes the chart type as one undoable step', () => {
		const { store, model } = setup();
		act(() => store.setUi({ selectedObject: 'chart-1', ribbonTab: 'chartDesign' }));
		fireEvent.click(screen.getByRole('button', { name: 'Change Chart Type' }));
		fireEvent.click(screen.getByRole('option', { name: 'Line with Markers' }));
		expect(model().plots[0].type).toBe('line');
		expect(model().dirty).toBe(true);
		expect(store.activeSheet.drawingsDirty).toBe(true);
		expect(store.getSnapshot().history.undo).toBe('Change Chart Type');
		act(() => store.undo());
		expect(model().plots[0].type).toBe('bar');
		act(() => store.redo());
		expect(model().plots[0].type).toBe('line');
	});

	it('adds chart elements, recolors and switches rows and columns', () => {
		const { store, model } = setup();
		act(() => store.setUi({ selectedObject: 'chart-1', ribbonTab: 'chartDesign' }));
		fireEvent.click(screen.getByRole('button', { name: 'Add Chart Element' }));
		fireEvent.click(screen.getByRole('menuitem', { name: 'Legend' }));
		fireEvent.click(screen.getByRole('menuitemcheckbox', { name: 'Right' }));
		expect(model().legend?.position).toBe('r');
		fireEvent.click(screen.getByRole('button', { name: 'Change Colors' }));
		fireEvent.click(screen.getByRole('option', { name: 'Colorful Palette 3' }));
		expect(model().plots[0].series[0].fill).toEqual({
			type: 'solid',
			color: { kind: 'scheme', value: 'accent2' },
		});
		fireEvent.click(screen.getByRole('button', { name: 'Switch Row/Column' }));
		expect(model().plots[0].series).toHaveLength(4);
		fireEvent.click(screen.getByRole('button', { name: 'Chart Styles' }));
		fireEvent.click(screen.getByRole('option', { name: 'Style 6' }));
		expect(model().chartFill).toMatchObject({ type: 'solid' });
		act(() => store.undo());
		act(() => store.undo());
		expect(model().plots[0].series).toHaveLength(2);
	});
});

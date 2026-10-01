import { act, fireEvent, render, screen } from '@testing-library/react';
import { createWorkbook, EditTransaction, groupLines } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import { MergeIndex } from '../../../engine/layout/style-at';
import { computePanes } from '../../../engine/layout/panes';
import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { GridRenderContextReact, type GridRenderContext } from './grid-context';
import { OutlineLevelButtons, RowOutlineGutter } from './Outline';
import { computeOutline } from 'xlsx-viewer-core';

function setup() {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	groupLines(tx, wb, 0, 'rows', 1, 3);
	tx.commit();
	const store = new ViewerStore(wb);
	const sheet = wb.sheets[0];
	const geometry = new SheetGeometry(sheet, { zoom: 1 });
	const ctx = {
		sheet,
		sheetIndex: 0,
		geometry,
		panes: computePanes(sheet, geometry),
		merges: new MergeIndex([]),
		styles: store.styles,
		colorCtx: { theme: wb.theme },
		zoom: 1,
		docVersion: 0,
		date1904: false,
		showGridLines: true,
		showFormulas: false,
		showZeros: true,
		gridlineColor: '#ddd',
		bodyRows: { start: 0, end: 20 },
		bodyCols: { start: 0, end: 10 },
	} as unknown as GridRenderContext;
	const renderGutter = () =>
		render(
			<ViewerStoreContext.Provider value={store}>
				<GridRenderContextReact.Provider value={ctx}>
					<RowOutlineGutter
						outline={computeOutline(sheet, 'rows')}
						span={{ start: 0, end: 20 }}
						size={32}
					/>
					<OutlineLevelButtons
						rows={computeOutline(sheet, 'rows')}
						cols={computeOutline(sheet, 'cols')}
						rowGutter={32}
						colGutter={0}
						headerHeight={20}
					/>
				</GridRenderContextReact.Provider>
			</ViewerStoreContext.Provider>,
		);
	return { store, sheet, renderGutter };
}

describe('outline gutter', () => {
	it('collapses a group with its - button and shows levels with the level buttons', () => {
		const { sheet, renderGutter } = setup();
		renderGutter();
		const toggle = screen.getByRole('button', { name: 'Hide Detail' });
		act(() => {
			fireEvent.click(toggle);
		});
		expect([1, 2, 3].every((r) => sheet.rows.get(r)?.hidden)).toBe(true);
		act(() => {
			fireEvent.click(screen.getByRole('button', { name: 'Show row level 2' }));
		});
		expect([1, 2, 3].some((r) => sheet.rows.get(r)?.hidden)).toBe(false);
		expect(screen.getAllByRole('button', { name: /Show row level/ })).toHaveLength(2);
	});
});

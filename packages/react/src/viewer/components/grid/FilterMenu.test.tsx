import { fireEvent, render, screen } from '@testing-library/react';
import { createWorkbook, putCell } from 'xlsx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { FilterMenu } from './FilterMenu';

/** Fruit / Apple / (blank) / Banana in A1:A4, with an AutoFilter on it. */
function setup() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	putCell(sheet, 0, 0, { value: 'Fruit' });
	putCell(sheet, 1, 0, { value: 'Apple' });
	putCell(sheet, 3, 0, { value: 'Banana' });
	sheet.autoFilter = { ref: { top: 0, left: 0, bottom: 3, right: 0 }, columns: [] };
	const store = new ViewerStore(wb);
	const onClose = vi.fn();
	render(
		<ViewerStoreContext.Provider value={store}>
			<FilterMenu target={{ kind: 'sheet' }} colId={0} onClose={onClose} />
		</ViewerStoreContext.Provider>,
	);
	return { store, sheet, onClose };
}

const box = (label: string) => screen.getByLabelText(label) as HTMLInputElement;

describe('FilterMenu', () => {
	it('lists the values and (Blanks), all checked', () => {
		setup();
		expect(screen.getByRole('dialog', { name: 'Filter' })).toBeTruthy();
		expect(['(Select All)', 'Apple', 'Banana', '(Blanks)'].map((l) => box(l).checked)).toEqual([
			true,
			true,
			true,
			true,
		]);
	});

	it('filters out blanks and records a Filter undo step', () => {
		const { store, sheet, onClose } = setup();
		fireEvent.click(box('(Blanks)'));
		expect(box('(Select All)').checked).toBe(false);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(sheet.autoFilter?.columns).toEqual([
			{ colId: 0, filters: { values: ['Apple', 'Banana'], blank: undefined } },
		]);
		expect(store.getSnapshot().history.undo).toBe('Filter');
	});

	it('searches the entry labels, including (Blanks)', () => {
		setup();
		fireEvent.change(screen.getByLabelText('Search values'), { target: { value: 'blan' } });
		expect(screen.queryByLabelText('Apple')).toBeNull();
		expect(box('(Blanks)').checked).toBe(true);
	});
});

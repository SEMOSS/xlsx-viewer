import { act, fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, getPrintArea, getPrintTitles } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog } from './__tests__/render-dialog';
import { PageSetupDialog } from './PageSetupDialog';

const tab = (name: string) => fireEvent.click(screen.getByRole('tab', { name }));
const type = (label: string, value: string) =>
	fireEvent.change(screen.getByLabelText(label), { target: { value } });

describe('PageSetupDialog', () => {
	it('applies every tab in one undoable step', () => {
		const store = new ViewerStore(createWorkbook());
		const { onClose } = renderDialog(PageSetupDialog, { store });
		fireEvent.click(screen.getByLabelText('Landscape'));
		type('Adjust to percent of normal size', '85');
		tab('Margins');
		type('Top:', '1.25');
		type('Left:', '0.5');
		fireEvent.click(screen.getByLabelText('Horizontally'));
		tab('Header/Footer');
		type('Footer center section', 'Page &P of &N');
		expect(screen.getByLabelText('Footer preview').textContent).toContain('Page 1 of 1');
		fireEvent.change(screen.getAllByRole('combobox')[0], { target: { value: '1' } });
		tab('Sheet');
		type('Print area:', 'A1:D20');
		type('Rows to repeat at top:', '$1:$2');
		fireEvent.click(screen.getByLabelText('Gridlines'));
		fireEvent.click(screen.getByLabelText('Over, then down'));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		const sheet = store.activeSheet;
		expect(sheet.pageSetup).toMatchObject({
			orientation: 'landscape',
			scale: 85,
			pageOrder: 'overThenDown',
		});
		expect(sheet.pageMargins).toMatchObject({ top: 1.25, left: 0.5 });
		expect(sheet.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
		expect(sheet.headerFooter).toEqual({ oddHeader: '&CPage &P', oddFooter: '&CPage &P of &N' });
		expect(getPrintArea(store.workbook, 0)).toEqual([{ top: 0, left: 0, bottom: 19, right: 3 }]);
		expect(getPrintTitles(store.workbook, 0).rows).toEqual({ start: 0, end: 1 });
		expect(store.getSnapshot().history.undo).toBe('Page Setup');
		act(() => store.undo());
		expect(sheet.pageSetup).toBeUndefined();
		expect(sheet.pageMargins).toBeUndefined();
		expect(store.workbook.definedNames).toEqual([]);
	});

	it('opens on the requested tab and fits to pages', () => {
		const store = new ViewerStore(createWorkbook());
		renderDialog(PageSetupDialog, { store, payload: { tab: 'margins' } });
		expect(screen.getByRole('tab', { name: 'Margins' }).getAttribute('aria-selected')).toBe('true');
		tab('Page');
		fireEvent.click(screen.getByLabelText('Fit to:'));
		type('Pages tall', '');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(store.activeSheet.pageSetup).toMatchObject({
			fitToPage: true,
			fitToWidth: 1,
			fitToHeight: 0,
		});
	});

	it('keeps the dialog open with an invalid entry', () => {
		const store = new ViewerStore(createWorkbook());
		const { onClose } = renderDialog(PageSetupDialog, { store, payload: { tab: 'sheet' } });
		type('Print area:', 'A1:(');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(screen.getByRole('alert').textContent).toBe("Reference isn't valid.");
		expect(onClose).not.toHaveBeenCalled();
		expect(store.getSnapshot().history.undo).toBeNull();
	});
});

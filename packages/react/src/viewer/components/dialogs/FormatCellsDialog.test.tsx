import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import { cellKey, createWorkbook, numberFormatCode, parseRange } from 'xlsx-viewer-core';

import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { FormatCellsDialog } from './FormatCellsDialog';

function xfAt(wb: Workbook, ref: string) {
	const r = parseRange(ref)!;
	return wb.styles.cellXfs[wb.sheets[0].cells.get(cellKey(r.top, r.left))?.style ?? 0];
}

function setup(tab?: string) {
	const wb = createWorkbook();
	typeCells(wb, { A1: '1234.5', A2: 'text' });
	const view = renderDialog(FormatCellsDialog, {
		workbook: wb,
		payload: tab ? { tab } : undefined,
	});
	return { ...view, wb };
}

function option(list: string, name: string) {
	return within(screen.getByRole('listbox', { name: list })).getByText(name);
}

describe('FormatCellsDialog', () => {
	it('opens on the requested tab and lists the six tabs', () => {
		setup('font');
		const tabs = screen.getAllByRole('tab').map((t) => t.textContent);
		expect(tabs).toEqual(['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection']);
		expect(screen.getByRole('tab', { name: 'Font' }).getAttribute('aria-selected')).toBe('true');
		expect(screen.getByTestId('font-preview').textContent).toBe('AaBbCcYyZz');
	});

	it('builds a number format with a live sample and applies it on OK', () => {
		const { wb, store, onClose } = setup('number');
		fireEvent.click(option('Category', 'Currency'));
		expect(screen.getByTestId('number-sample').textContent).toBe('$1,234.50');
		fireEvent.change(screen.getByRole('spinbutton', { name: 'Decimal places:' }), {
			target: { value: '0' },
		});
		expect(screen.getByTestId('number-sample').textContent).toBe('$1,235');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(numberFormatCode(xfAt(wb, 'A1').numFmtId, wb.styles.numFmts)).toBe('"$"#,##0');
		expect(store.getSnapshot().history.undo).toBe('Format Cells');
	});

	it('edits a custom code', () => {
		const { wb } = setup('number');
		fireEvent.click(option('Category', 'Custom'));
		const input = screen.getByRole('textbox', { name: 'Type:' });
		fireEvent.change(input, { target: { value: '0.0,,"M"' } });
		fireEvent.keyDown(input, { key: 'Enter' });
		expect(numberFormatCode(xfAt(wb, 'A1').numFmtId, wb.styles.numFmts)).toBe('0.0,,"M"');
	});

	it('applies font, alignment, fill and protection changes together', () => {
		const { wb, store } = setup('font');
		selectRange(store, 'A1:A2');
		fireEvent.click(option('Font style', 'Bold Italic'));
		fireEvent.click(screen.getByRole('checkbox', { name: 'Strikethrough' }));
		fireEvent.click(screen.getByRole('tab', { name: 'Alignment' }));
		fireEvent.change(screen.getByRole('combobox', { name: 'Horizontal:' }), {
			target: { value: 'center' },
		});
		fireEvent.click(screen.getByRole('checkbox', { name: 'Wrap text' }));
		fireEvent.click(screen.getByRole('tab', { name: 'Protection' }));
		fireEvent.click(screen.getByRole('checkbox', { name: 'Locked' }));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		for (const ref of ['A1', 'A2']) {
			const xf = xfAt(wb, ref);
			const font = wb.styles.fonts[xf.fontId];
			expect(font).toMatchObject({ bold: true, italic: true, strike: true, name: 'Calibri' });
			expect(xf.alignment).toEqual({ horizontal: 'center', wrapText: true });
			expect(xf.protection).toEqual({ locked: false });
			expect(numberFormatCode(xf.numFmtId, wb.styles.numFmts)).toBe('General');
		}
	});

	it('draws an outline from the Border tab presets', () => {
		const { wb, store } = setup('border');
		selectRange(store, 'A1:B2');
		fireEvent.click(screen.getByRole('radio', { name: 'thick' }));
		fireEvent.click(screen.getByRole('button', { name: 'Outline' }));
		expect(screen.getByRole('button', { name: 'Top border' }).getAttribute('aria-pressed')).toBe(
			'true',
		);
		fireEvent.click(screen.getByRole('button', { name: 'Diagonal down border' }));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		const border = (ref: string) => wb.styles.borders[xfAt(wb, ref).borderId];
		expect(border('A1')).toMatchObject({
			top: { style: 'thick' },
			left: { style: 'thick' },
			diagonalDown: true,
		});
		expect(border('B2')).toMatchObject({ bottom: { style: 'thick' }, right: { style: 'thick' } });
		expect(border('B2').top).toBeUndefined();
	});

	it('fills with a background color', () => {
		const { wb } = setup('fill');
		fireEvent.click(screen.getByRole('button', { name: '#FFFF00' }));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(wb.styles.fills[xfAt(wb, 'A1').fillId]).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { rgb: 'FFFFFF00' },
			bgColor: { indexed: 64 },
		});
	});

	it('changes nothing on Escape', () => {
		const { store, onClose } = setup('alignment');
		fireEvent.click(screen.getByRole('checkbox', { name: 'Wrap text' }));
		fireEvent.keyDown(screen.getByRole('checkbox', { name: 'Wrap text' }), { key: 'Escape' });
		expect(onClose).toHaveBeenCalledTimes(1);
		expect(store.isDirty).toBe(false);
	});

	it('adds no undo step when OK is pressed without edits', () => {
		const { store, onClose } = setup('number');
		fireEvent.click(screen.getByRole('tab', { name: 'Border' }));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.getSnapshot().history.undo).toBeNull();
	});

	it('sets the text angle with the orientation dial', () => {
		const { wb } = setup('alignment');
		const dial = screen.getByRole('slider', { name: 'Text orientation' });
		fireEvent.keyDown(dial, { key: 'PageUp' });
		fireEvent.keyDown(dial, { key: 'PageUp' });
		fireEvent.keyDown(dial, { key: 'PageUp' });
		expect((screen.getByRole('spinbutton', { name: 'Degrees' }) as HTMLInputElement).value).toBe(
			'45',
		);
		fireEvent.change(screen.getByRole('spinbutton', { name: 'Degrees' }), {
			target: { value: '-30' },
		});
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(xfAt(wb, 'A1').alignment).toEqual({ textRotation: 120 });
	});
});

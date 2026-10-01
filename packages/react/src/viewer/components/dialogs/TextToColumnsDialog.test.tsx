import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import { createWorkbook, getCell, numberFormatCode, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { TextToColumnsDialog } from './TextToColumnsDialog';

function open(entries: Record<string, string>, ref: string, readOnly = false) {
	const wb = createWorkbook();
	typeCells(wb, entries);
	const store = new ViewerStore(wb, { readOnly });
	selectRange(store, ref);
	return { wb, ...renderDialog(TextToColumnsDialog, { store }) };
}

const value = (wb: Workbook, ref: string) => {
	const r = parseRange(ref)!;
	return getCell(wb.sheets[0], r.top, r.left)?.value ?? null;
};
const button = (name: string) => screen.getByRole('button', { name }) as HTMLButtonElement;
const title = () => screen.getByRole('dialog').getAttribute('aria-label');
const check = (name: string) => screen.getByRole('checkbox', { name }) as HTMLInputElement;
const radio = (name: string) => screen.getByRole('radio', { name }) as HTMLInputElement;

describe('TextToColumnsDialog', () => {
	it('explains that only one column can be converted', () => {
		const { store, onClose, wb } = open({ A1: 'a,b', B1: 'c' }, 'A1:B1');
		const message =
			'Microsoft Excel can convert only one column at a time. The range can be many rows tall ' +
			'but no more than one column wide. Try again by selecting cells in one column only.';
		expect(screen.getByRole('dialog', { name: 'Microsoft Excel' })).toBeTruthy();
		expect(screen.getByRole('alert').textContent).toBe(message);
		expect(store.state.message).toBe(message);
		fireEvent.click(button('OK'));
		expect(onClose).toHaveBeenCalled();
		expect(value(wb, 'A1')).toBe('a,b');
	});

	it('reports an empty selection', () => {
		open({ B1: 'x' }, 'A1:A3');
		expect(screen.getByRole('alert').textContent).toBe('No data was selected to parse.');
	});

	it('walks through the three steps and converts delimited text', () => {
		const { wb, store, onClose } = open({ A1: 'Ann,007', A2: 'Bob,8', A3: 'Cid' }, 'A1:A3');
		expect(title()).toBe('Convert Text to Columns Wizard - Step 1 of 3');
		expect(screen.getByText(/has determined that your data is Delimited/)).toBeTruthy();
		expect(radio('Delimited').checked).toBe(true);
		expect(screen.getByRole('region', { name: 'Preview of selected data:' }).textContent).toContain(
			'Ann,007',
		);
		expect(button('< Back').disabled).toBe(true);

		fireEvent.click(button('Next >'));
		expect(title()).toBe('Convert Text to Columns Wizard - Step 2 of 3');
		expect(check('Tab').checked).toBe(true);
		fireEvent.click(check('Comma'));
		const preview = screen.getByRole('table', { name: 'Data preview' });
		expect(within(preview).getAllByRole('row')[0].textContent).toBe('Ann007');

		fireEvent.click(button('Next >'));
		expect(title()).toBe('Convert Text to Columns Wizard - Step 3 of 3');
		expect(button('Next >').disabled).toBe(true);
		expect((screen.getByRole('textbox', { name: 'Destination:' }) as HTMLInputElement).value).toBe(
			'$A$1',
		);
		fireEvent.click(button('Column 2: General'));
		fireEvent.click(radio('Text'));
		expect(button('Column 2: Text').getAttribute('aria-pressed')).toBe('true');

		fireEvent.click(button('Finish'));
		expect(onClose).toHaveBeenCalled();
		expect([
			value(wb, 'A1'),
			value(wb, 'B1'),
			value(wb, 'A2'),
			value(wb, 'B2'),
			value(wb, 'A3'),
		]).toEqual(['Ann', '007', 'Bob', '8', 'Cid']);
		const b1 = getCell(wb.sheets[0], 0, 1);
		expect(numberFormatCode(wb.styles.cellXfs[b1?.style ?? 0].numFmtId, wb.styles.numFmts)).toBe(
			'@',
		);
		expect(store.selectionOf().ranges).toEqual([parseRange('A1:B3')]);
		expect(store.getSnapshot().history.undo).toBe('Text to Columns');
	});

	it('moves with Enter and Back, skips columns and honours the destination', () => {
		const { wb } = open({ A1: 'x;1;y' }, 'A1');
		const first = radio('Delimited');
		fireEvent.keyDown(first, { key: 'Enter' });
		fireEvent.click(check('Tab'));
		fireEvent.click(check('Semicolon'));
		fireEvent.click(button('< Back'));
		expect(title()).toContain('Step 1 of 3');
		fireEvent.keyDown(radio('Delimited'), { key: 'Enter' });
		expect(check('Semicolon').checked).toBe(true);
		fireEvent.keyDown(check('Semicolon'), { key: 'Enter' });
		fireEvent.click(radio('Do not import column (skip)'));
		expect(button('Column 1: Skip column')).toBeTruthy();
		const dest = screen.getByRole('textbox', { name: 'Destination:' });
		fireEvent.change(dest, { target: { value: 'C2' } });
		fireEvent.keyDown(dest, { key: 'Enter' });
		expect([value(wb, 'C2'), value(wb, 'D2'), value(wb, 'A1')]).toEqual([1, 'y', 'x;1;y']);
	});

	it('asks before replacing data and lets Escape answer Cancel', () => {
		const { wb, onClose } = open({ A1: 'a,b', B1: 'old' }, 'A1');
		fireEvent.click(button('Next >'));
		fireEvent.click(check('Comma'));
		fireEvent.click(button('Finish'));
		const ask = screen.getByRole('alertdialog', { name: 'Microsoft Excel' });
		expect(ask.textContent).toContain("There's already data here. Do you want to replace it?");
		fireEvent.keyDown(within(ask).getByRole('button', { name: 'OK' }), { key: 'Escape' });
		expect(onClose).not.toHaveBeenCalled();
		expect(screen.queryByRole('alertdialog')).toBeNull();
		expect(value(wb, 'B1')).toBe('old');

		fireEvent.click(button('Finish'));
		fireEvent.click(
			within(screen.getByRole('alertdialog')).getByRole('button', { name: 'Cancel' }),
		);
		expect(value(wb, 'B1')).toBe('old');
		fireEvent.click(button('Finish'));
		fireEvent.click(within(screen.getByRole('alertdialog')).getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect([value(wb, 'A1'), value(wb, 'B1')]).toEqual(['a', 'b']);
	});

	it('guesses fixed width and edits break lines', () => {
		const { wb } = open({ A1: 'Ann   42', A2: 'Bob   7' }, 'A1:A2');
		expect(radio('Fixed width').checked).toBe(true);
		fireEvent.click(button('Next >'));
		const breaks = screen.getByRole('textbox', { name: 'Break positions:' }) as HTMLInputElement;
		expect(breaks.value).toBe('6');
		const area = screen.getByTestId('fixed-width-area');
		fireEvent.click(area, { clientX: 3 * 8 });
		expect(breaks.value).toBe('3, 6');
		// Drag the break at 6 to 5 (the test layout uses 8px characters).
		fireEvent.mouseDown(area.querySelector('[data-break="6"]')!, { clientX: 6 * 8 });
		fireEvent.mouseMove(window, { clientX: 5 * 8 });
		fireEvent.mouseUp(window, { clientX: 5 * 8 });
		expect(breaks.value).toBe('3, 5');
		fireEvent.doubleClick(area.querySelector('[data-break="5"]')!);
		expect(breaks.value).toBe('3');
		fireEvent.change(breaks, { target: { value: '2, 6' } });
		fireEvent.blur(breaks);
		expect(breaks.value).toBe('2, 6');
		fireEvent.click(button('Finish'));
		expect([value(wb, 'A1'), value(wb, 'B1'), value(wb, 'C1')]).toEqual(['An', 'n', 42]);
		expect([value(wb, 'A2'), value(wb, 'B2'), value(wb, 'C2')]).toEqual(['Bo', 'b', 7]);
	});

	it('rejects an invalid destination and cancels without changes', () => {
		const { wb, store, onClose } = open({ A1: 'a\tb' }, 'A1');
		fireEvent.click(button('Next >'));
		fireEvent.click(button('Next >'));
		fireEvent.change(screen.getByRole('textbox', { name: 'Destination:' }), {
			target: { value: 'nope!' },
		});
		fireEvent.click(button('Finish'));
		expect(screen.getByRole('alert').textContent).toBe("Reference isn't valid.");
		expect(onClose).not.toHaveBeenCalled();
		fireEvent.keyDown(radio('General'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
		expect(value(wb, 'A1')).toBe('a\tb');
		expect(store.isDirty).toBe(false);
	});

	it('disables Finish when read-only', () => {
		open({ A1: 'a,b' }, 'A1', true);
		expect(button('Finish').disabled).toBe(true);
	});
});

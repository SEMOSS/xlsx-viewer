import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { DataValidation } from 'xlsx-viewer-core';
import { createWorkbook, formatRange, parseRange } from 'xlsx-viewer-core';

import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange } from './__tests__/render-dialog';
import { DataValidationDialog } from './DataValidationDialog';

const r = (a1: string) => parseRange(a1)!;
const refs = (dv: DataValidation) => dv.ranges.map((x) => formatRange(x));

function storeWith(list: DataValidation[], select: string, readOnly = false) {
	const wb = createWorkbook();
	wb.sheets[0].dataValidations = list;
	const store = new ViewerStore(wb, { readOnly });
	selectRange(store, select);
	return store;
}

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;
const change = (label: string, value: string) =>
	fireEvent.change(field(label), { target: { value } });
const click = (name: string) => fireEvent.click(screen.getByRole('button', { name }));
const validations = (store: ViewerStore) => store.workbook.sheets[0].dataValidations;

describe('DataValidationDialog', () => {
	it('creates a whole-number rule for the selection', () => {
		const store = storeWith([], 'A1:A5');
		const { onClose } = renderDialog(DataValidationDialog, { store });
		expect(screen.getByRole('tab', { name: 'Settings' }).getAttribute('aria-selected')).toBe(
			'true',
		);
		expect(field('Allow:').value).toBe('none');
		change('Allow:', 'whole');
		expect(field('Data:').value).toBe('between');
		change('Minimum:', '1');
		change('Maximum:', '10');
		click('OK');
		expect(onClose).toHaveBeenCalled();
		expect(validations(store)).toEqual([
			{
				type: 'whole',
				operator: 'between',
				formula1: '1',
				formula2: '10',
				allowBlank: true,
				showInputMessage: true,
				showErrorMessage: true,
				errorStyle: 'stop',
				ranges: [r('A1:A5')],
			},
		]);
		expect(store.getSnapshot().history.undo).toBe('Data Validation');
	});

	it('shows Excel messages inline and keeps the dialog open', () => {
		const store = storeWith([], 'B2');
		const { onClose } = renderDialog(DataValidationDialog, { store });
		change('Allow:', 'decimal');
		change('Maximum:', '5');
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'You must enter a value in the Minimum box.',
		);
		change('Minimum:', '10');
		expect(screen.queryByRole('alert')).toBeNull();
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'The Maximum must be greater than or equal to the Minimum.',
		);
		change('Allow:', 'list');
		click('OK');
		expect(screen.getByRole('alert').textContent).toBe(
			'The list source must be a delimited list, or a reference to single row or column.',
		);
		expect(onClose).not.toHaveBeenCalled();
		expect(validations(store)).toEqual([]);
		expect(store.isDirty).toBe(false);
	});

	it('labels the value boxes like Excel for each type and operator', () => {
		renderDialog(DataValidationDialog, { store: storeWith([], 'A1') });
		expect(screen.queryByLabelText('Data:')).toBeNull();
		expect((screen.getByLabelText('Ignore blank') as HTMLInputElement).disabled).toBe(true);
		change('Allow:', 'date');
		change('Data:', 'greaterThan');
		expect(screen.getByLabelText('Start date:')).toBeTruthy();
		expect(screen.queryByLabelText('End date:')).toBeNull();
		change('Data:', 'equal');
		expect(screen.getByLabelText('Date:')).toBeTruthy();
		change('Allow:', 'time');
		// The Data choice is kept when the type changes.
		expect(screen.getByLabelText('Time:')).toBeTruthy();
		change('Data:', 'notBetween');
		expect(screen.getByLabelText('Start time:')).toBeTruthy();
		expect(screen.getByLabelText('End time:')).toBeTruthy();
		change('Allow:', 'textLength');
		change('Data:', 'notEqual');
		expect(screen.getByLabelText('Length:')).toBeTruthy();
		change('Allow:', 'list');
		expect(screen.queryByLabelText('Data:')).toBeNull();
		expect(screen.getByLabelText('Source:')).toBeTruthy();
		expect((screen.getByLabelText('In-cell dropdown') as HTMLInputElement).checked).toBe(true);
		change('Allow:', 'custom');
		expect(screen.getByLabelText('Formula:')).toBeTruthy();
		expect(screen.queryByLabelText('In-cell dropdown')).toBeNull();
	});

	it('creates a list with an input message and a warning alert, OK on Enter', () => {
		const store = storeWith([], 'C1:C3');
		const { onClose } = renderDialog(DataValidationDialog, { store });
		change('Allow:', 'list');
		change('Source:', 'Yes,No');
		fireEvent.click(screen.getByLabelText('In-cell dropdown'));
		fireEvent.click(screen.getByRole('tab', { name: 'Input Message' }));
		expect(
			(screen.getByLabelText('Show input message when cell is selected') as HTMLInputElement)
				.checked,
		).toBe(true);
		change('Title:', 'Choose');
		change('Input message:', 'Yes or no');
		fireEvent.click(screen.getByRole('tab', { name: 'Error Alert' }));
		expect(screen.getByRole('img', { name: 'Stop icon' })).toBeTruthy();
		change('Style:', 'warning');
		expect(screen.getByRole('img', { name: 'Warning icon' })).toBeTruthy();
		change('Error message:', 'Please pick one');
		change('Title:', 'Hmm');
		fireEvent.keyDown(field('Title:'), { key: 'Enter' });
		expect(onClose).toHaveBeenCalled();
		const [dv] = validations(store);
		expect(dv).toMatchObject({
			type: 'list',
			formula1: '"Yes,No"',
			inCellDropdown: false,
			allowBlank: true,
			promptTitle: 'Choose',
			prompt: 'Yes or no',
			errorStyle: 'warning',
			errorTitle: 'Hmm',
			error: 'Please pick one',
			ranges: [r('C1:C3')],
		});
		expect(dv.operator).toBeUndefined();
	});

	it('prefills from the selection and changes only the selected cells', () => {
		const existing: DataValidation = {
			ranges: [r('A1:A10')],
			type: 'date',
			operator: 'between',
			formula1: '45292',
			formula2: '45657',
			allowBlank: true,
			showErrorMessage: true,
		};
		const store = storeWith([existing], 'A2:A3');
		renderDialog(DataValidationDialog, { store });
		expect(field('Allow:').value).toBe('date');
		expect(field('Start date:').value).toBe('1/1/2024');
		expect(field('End date:').value).toBe('12/31/2024');
		fireEvent.click(screen.getByRole('tab', { name: 'Input Message' }));
		expect(
			(screen.getByLabelText('Show input message when cell is selected') as HTMLInputElement)
				.checked,
		).toBe(false);
		fireEvent.click(screen.getByRole('tab', { name: 'Settings' }));
		change('End date:', '6/30/2024');
		click('OK');
		const list = validations(store);
		expect(list.map(refs)).toEqual([['A1', 'A4:A10'], ['A2:A3']]);
		expect(list[0].formula2).toBe('45657');
		expect(list[1]).toMatchObject({ type: 'date', formula1: '45292', formula2: '45473' });
	});

	it('Clear All resets the fields and OK then removes the validation', () => {
		const store = storeWith(
			[{ ranges: [r('A1:A3')], type: 'list', formula1: '"a,b"', allowBlank: false }],
			'A1:A3',
		);
		const { onClose } = renderDialog(DataValidationDialog, { store });
		expect(field('Source:').value).toBe('a,b');
		expect((screen.getByLabelText('Ignore blank') as HTMLInputElement).checked).toBe(false);
		click('Clear All');
		expect(onClose).not.toHaveBeenCalled();
		expect(field('Allow:').value).toBe('none');
		expect(validations(store)).toHaveLength(1);
		click('OK');
		expect(validations(store)).toEqual([]);
		store.undo();
		expect(validations(store)).toHaveLength(1);
	});

	it('asks before changing a selection with several validations', () => {
		const list: DataValidation[] = [
			{ ranges: [r('A1')], type: 'whole', operator: 'greaterThan', formula1: '0' },
			{ ranges: [r('A2')], type: 'list', formula1: '"x"' },
		];
		const first = renderDialog(DataValidationDialog, { store: storeWith(list, 'A1:A2') });
		expect(screen.getByText(/more than one type of validation/).textContent).toBe(
			'The selection contains more than one type of validation. Erase current settings and continue?',
		);
		click('No');
		expect(first.onClose).toHaveBeenCalled();
		first.unmount();

		const store = storeWith(list, 'A1:A2');
		renderDialog(DataValidationDialog, { store });
		click('Yes');
		expect(field('Allow:').value).toBe('none');
		click('OK');
		expect(validations(store)).toEqual([]);
	});

	it('asks before extending a validation to cells without one', () => {
		const list: DataValidation[] = [
			{ ranges: [r('A1')], type: 'list', formula1: '"a,b"', allowBlank: true },
		];
		const yes = storeWith(list, 'A1:A3');
		const first = renderDialog(DataValidationDialog, { store: yes });
		expect(screen.getByText(/without Data Validation settings/).textContent).toBe(
			'The selection contains some cells without Data Validation settings. Do you want to extend Data Validation to these cells?',
		);
		click('Yes');
		expect(field('Source:').value).toBe('a,b');
		click('OK');
		expect(validations(yes).map(refs)).toEqual([['A1:A3']]);
		first.unmount();

		const no = storeWith(list, 'A1:A3');
		const second = renderDialog(DataValidationDialog, { store: no });
		click('No');
		change('Source:', 'c,d');
		click('OK');
		expect(validations(no).map(refs)).toEqual([['A1']]);
		expect(validations(no)[0].formula1).toBe('"c,d"');
		second.unmount();

		const { onClose } = renderDialog(DataValidationDialog, { store: storeWith(list, 'A1:A3') });
		click('Cancel');
		expect(onClose).toHaveBeenCalled();
	});

	it('applies the changes to all other cells with the same settings', () => {
		const rule = { type: 'list' as const, formula1: '"a,b"', allowBlank: true };
		const store = storeWith(
			[
				{ ...rule, ranges: [r('A1')] },
				{ ...rule, ranges: [r('C5')] },
			],
			'A1',
		);
		renderDialog(DataValidationDialog, { store });
		const applyAll = screen.getByLabelText(
			'Apply these changes to all other cells with the same settings',
		) as HTMLInputElement;
		expect(applyAll.disabled).toBe(false);
		fireEvent.click(applyAll);
		change('Source:', 'x,y');
		click('OK');
		expect(validations(store)).toHaveLength(1);
		expect(validations(store)[0]).toMatchObject({ formula1: '"x,y"' });
		expect(refs(validations(store)[0])).toEqual(['A1', 'C5']);
	});

	it('disables Apply to all without an existing validation', () => {
		renderDialog(DataValidationDialog, { store: storeWith([], 'A1') });
		const applyAll = screen.getByLabelText(
			'Apply these changes to all other cells with the same settings',
		) as HTMLInputElement;
		expect(applyAll.disabled).toBe(true);
	});

	it('cancels with Escape or Cancel without changing anything', () => {
		const store = storeWith([], 'A1');
		const first = renderDialog(DataValidationDialog, { store });
		change('Allow:', 'list');
		change('Source:', 'a,b');
		fireEvent.keyDown(field('Source:'), { key: 'Escape' });
		expect(first.onClose).toHaveBeenCalled();
		first.unmount();
		const second = renderDialog(DataValidationDialog, { store });
		change('Allow:', 'whole');
		click('Cancel');
		expect(second.onClose).toHaveBeenCalled();
		expect(validations(store)).toEqual([]);
		expect(store.isDirty).toBe(false);
	});

	it('disables OK and Clear All on a read-only workbook', () => {
		const store = storeWith([{ ranges: [r('A1')], type: 'list', formula1: '"a"' }], 'A1', true);
		renderDialog(DataValidationDialog, { store });
		expect((screen.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true);
		expect((screen.getByRole('button', { name: 'Clear All' }) as HTMLButtonElement).disabled).toBe(
			true,
		);
		fireEvent.keyDown(field('Source:'), { key: 'Enter' });
		expect(validations(store)[0].formula1).toBe('"a"');
	});
});

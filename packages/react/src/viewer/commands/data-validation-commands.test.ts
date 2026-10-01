import { describe, expect, it } from 'vitest';
import type { DataValidation, DataValidationOperator, DataValidationType } from 'xlsx-viewer-core';
import { createWorkbook, formatRange, parseRange } from 'xlsx-viewer-core';

import { selectRange } from '../components/dialogs/__tests__/render-dialog';
import { createTranslator, type Translate } from '../../i18n';
import { ViewerStore } from '../state/viewer-store';
import {
	applyValidationForm,
	dataValidationSetup,
	DEFAULT_VALIDATION_FORM,
	ruleFromValidationForm,
	validationFormFromRule,
	validationValueCaptions,
	type ValidationForm,
} from './data-validation-commands';

const r = (a1: string) => parseRange(a1)!;

function storeWith(...list: DataValidation[]) {
	const wb = createWorkbook({ sheetNames: ['Sheet1', 'Lists'] });
	wb.sheets[0].dataValidations = list;
	return new ViewerStore(wb);
}

const form = (patch: Partial<ValidationForm>): ValidationForm => ({
	...DEFAULT_VALIDATION_FORM,
	...patch,
});

describe('validationFormFromRule', () => {
	it('prefills every field from the rule', () => {
		const rule: DataValidation = {
			ranges: [r('A1:A9')],
			type: 'date',
			operator: 'notBetween',
			formula1: '45292',
			formula2: '45657',
			allowBlank: true,
			showInputMessage: true,
			promptTitle: 'When',
			prompt: 'Pick a 2024 date',
			showErrorMessage: true,
			errorStyle: 'warning',
			errorTitle: 'Oops',
			error: 'Not in 2024',
		};
		expect(validationFormFromRule(rule, { row: 0, col: 0 }, false)).toEqual({
			type: 'date',
			operator: 'notBetween',
			value1: '1/1/2024',
			value2: '12/31/2024',
			ignoreBlank: true,
			inCellDropdown: true,
			showInputMessage: true,
			promptTitle: 'When',
			prompt: 'Pick a 2024 date',
			showErrorMessage: true,
			errorStyle: 'warning',
			errorTitle: 'Oops',
			error: 'Not in 2024',
		});
	});

	it('uses file defaults for missing flags and shows relative references from the active cell', () => {
		const rule: DataValidation = { ranges: [r('B2:B10')], type: 'custom', formula1: 'C2>$A$1' };
		const f = validationFormFromRule(rule, { row: 4, col: 1 }, false);
		expect(f.value1).toBe('=C5>$A$1');
		expect(f.operator).toBe('between');
		expect(f.ignoreBlank).toBe(false);
		expect(f.showInputMessage).toBe(false);
		expect(f.showErrorMessage).toBe(false);
		expect(f.errorStyle).toBe('stop');
		const list = validationFormFromRule(
			{ ranges: [r('A1')], type: 'list', formula1: '"a,b"', inCellDropdown: false },
			{ row: 0, col: 0 },
			false,
		);
		expect(list).toMatchObject({ value1: 'a,b', inCellDropdown: false });
	});
});

describe('ruleFromValidationForm', () => {
	it('removes validation for Any value without an input message', () => {
		expect(ruleFromValidationForm(DEFAULT_VALIDATION_FORM, false)).toBeNull();
		expect(ruleFromValidationForm(form({ errorTitle: 'x' }), false)).toBeNull();
		expect(ruleFromValidationForm(form({ prompt: 'Type your name' }), false)).toEqual({
			type: 'none',
			allowBlank: true,
			showInputMessage: true,
			showErrorMessage: true,
			errorStyle: 'stop',
			prompt: 'Type your name',
		});
	});

	it('builds operator rules with the formulas the boxes need', () => {
		expect(
			ruleFromValidationForm(
				form({ type: 'whole', operator: 'between', value1: '1', value2: '10' }),
				false,
			),
		).toEqual({
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			allowBlank: true,
			showInputMessage: true,
			showErrorMessage: true,
			errorStyle: 'stop',
		});
		const single = ruleFromValidationForm(
			form({ type: 'decimal', operator: 'greaterThan', value1: '0.5', value2: '9' }),
			false,
		);
		expect(single?.formula1).toBe('0.5');
		expect(single?.formula2).toBeUndefined();
	});

	it('builds list rules with the dropdown flag and no operator', () => {
		const rule = ruleFromValidationForm(
			form({ type: 'list', value1: 'Yes,No', inCellDropdown: false, ignoreBlank: false }),
			false,
		);
		expect(rule).toMatchObject({
			type: 'list',
			formula1: '"Yes,No"',
			inCellDropdown: false,
			allowBlank: false,
		});
		expect(rule?.operator).toBeUndefined();
		expect(rule?.fromExtension).toBeUndefined();
	});

	it('marks rules referring to other sheets for the 2010 extension', () => {
		expect(
			ruleFromValidationForm(form({ type: 'list', value1: '=Lists!$A$1:$A$3' }), false)
				?.fromExtension,
		).toBe(true);
		expect(
			ruleFromValidationForm(form({ type: 'list', value1: 'Hi!,Bye' }), false)?.fromExtension,
		).toBeUndefined();
	});

	it('keeps messages as typed and drops blank ones', () => {
		const rule = ruleFromValidationForm(
			form({
				type: 'textLength',
				operator: 'lessThan',
				value1: '5',
				errorTitle: '  ',
				error: 'Too long ',
				promptTitle: 'Hint',
			}),
			false,
		);
		expect(rule).toMatchObject({ error: 'Too long ', promptTitle: 'Hint' });
		expect('errorTitle' in (rule ?? {})).toBe(false);
	});
});

describe('validationValueCaptions', () => {
	it('captions the value boxes per type and operator, in the UI language', () => {
		const t = createTranslator();
		expect(validationValueCaptions('whole', 'between', t)).toEqual(['Minimum:', 'Maximum:']);
		expect(validationValueCaptions('date', 'greaterThan', t)).toEqual(['Start date:']);
		expect(validationValueCaptions('time', 'lessThan', t)).toEqual(['End time:']);
		expect(validationValueCaptions('textLength', 'equal', t)).toEqual(['Length:']);
		expect(validationValueCaptions('list', 'between', t)).toEqual(['Source:']);
		expect(validationValueCaptions('custom', 'between', t)).toEqual(['Formula:']);
		expect(validationValueCaptions('none', 'between', t)).toEqual([]);
		const de = createTranslator({ 'dialog.dataValidation.box.source': 'Quelle:' });
		expect(validationValueCaptions('list', 'between', de)).toEqual(['Quelle:']);
	});

	it('has a key for every value box label core gives', () => {
		const keyOf: Translate = (key) => key;
		const types: DataValidationType[] = [
			'none',
			'whole',
			'decimal',
			'list',
			'date',
			'time',
			'textLength',
			'custom',
		];
		const operators: DataValidationOperator[] = [
			'between',
			'notBetween',
			'equal',
			'notEqual',
			'greaterThan',
			'lessThan',
			'greaterThanOrEqual',
			'lessThanOrEqual',
		];
		for (const type of types) {
			for (const operator of operators) {
				for (const caption of validationValueCaptions(type, operator, keyOf)) {
					expect(caption).toMatch(/^dialog\.dataValidation\.box\./);
				}
			}
		}
	});
});

describe('applyValidationForm', () => {
	it('returns Excel messages without changing anything', () => {
		const store = storeWith();
		selectRange(store, 'A1:A3');
		expect(
			applyValidationForm(store, form({ type: 'whole', value1: '', value2: '5' }), [r('A1:A3')]),
		).toBe('You must enter a value in the Minimum box.');
		expect(store.isDirty).toBe(false);
		expect(store.workbook.sheets[0].dataValidations).toEqual([]);
	});

	it('applies the rule to the targets in one undoable step', () => {
		const store = storeWith();
		selectRange(store, 'B2:B6');
		const sel = store.selectionOf();
		store.setSelection({ ...sel, active: { row: 3, col: 1 } });
		expect(
			applyValidationForm(store, form({ type: 'custom', value1: '=C4>0' }), [r('B2:B6')]),
		).toBeNull();
		const [dv] = store.workbook.sheets[0].dataValidations;
		expect(dv).toMatchObject({ type: 'custom', formula1: 'C2>0', ranges: [r('B2:B6')] });
		expect(store.getSnapshot().history.undo).toBe('Data Validation');
		store.undo();
		expect(store.workbook.sheets[0].dataValidations).toEqual([]);
		store.redo();
		expect(store.workbook.sheets[0].dataValidations).toHaveLength(1);
	});

	it('changes every cell with the same settings when asked to', () => {
		const base = {
			type: 'whole' as const,
			operator: 'between' as const,
			formula1: '1',
			formula2: '5',
		};
		const store = storeWith(
			{ ...base, ranges: [r('A1:A2')] },
			{ ...base, ranges: [r('D4')] },
			{ ranges: [r('F1')], type: 'list', formula1: '"x"' },
		);
		selectRange(store, 'A1');
		const setup = dataValidationSetup(store);
		expect(setup.state.rule?.type).toBe('whole');
		applyValidationForm(store, { ...setup.form, value2: '50' }, setup.selection, setup.state.rule);
		const list = store.workbook.sheets[0].dataValidations;
		expect(list).toHaveLength(2);
		expect(list[0].type).toBe('list');
		expect(list[1].formula2).toBe('50');
		expect(list[1].ranges.map((x) => formatRange(x))).toEqual(['A1', 'A2', 'D4']);
	});

	it('does nothing on a read-only store', () => {
		const wb = createWorkbook();
		const store = new ViewerStore(wb, { readOnly: true });
		expect(applyValidationForm(store, form({ type: 'list', value1: 'a,b' }), [r('A1')])).toBeNull();
		expect(wb.sheets[0].dataValidations).toEqual([]);
	});
});

describe('dataValidationSetup', () => {
	it('starts from defaults without validation and from the rule otherwise', () => {
		const store = storeWith({
			ranges: [r('C1:C5')],
			type: 'list',
			formula1: '"Red,Green"',
			allowBlank: true,
		});
		selectRange(store, 'A1');
		expect(dataValidationSetup(store).form).toEqual(DEFAULT_VALIDATION_FORM);
		selectRange(store, 'C2:C3');
		const setup = dataValidationSetup(store);
		expect(setup.form).toMatchObject({ type: 'list', value1: 'Red,Green', ignoreBlank: true });
		expect(setup.state).toMatchObject({ mixed: false, partial: false });
		selectRange(store, 'C4:C8');
		expect(dataValidationSetup(store).state.partial).toBe(true);
	});
});

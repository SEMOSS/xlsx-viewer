import { describe, expect, it } from 'vitest';

import { formatRange, parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import type { DataValidation } from '../types/data-validation';
import { findValidation } from '../validation/find';
import {
	EditTransaction,
	sameValidationRanges,
	sameValidationSettings,
	selectionValidationState,
	setDataValidation,
	type DataValidationRule,
} from './index';

const r = (a1: string) => parseRange(a1)!;
const refs = (dv: { ranges: DataValidation['ranges'] }) => dv.ranges.map((x) => formatRange(x));

const WHOLE: DataValidationRule = {
	type: 'whole',
	operator: 'between',
	formula1: '1',
	formula2: '10',
	allowBlank: true,
	showInputMessage: true,
	showErrorMessage: true,
};

const LIST: DataValidationRule = {
	type: 'list',
	formula1: '"Yes,No"',
	allowBlank: true,
	inCellDropdown: true,
	showInputMessage: true,
	showErrorMessage: true,
};

function withValidations(...list: DataValidation[]) {
	const wb = createWorkbook();
	wb.sheets[0].dataValidations = list;
	wb.sheets[0].dirty = false;
	return wb;
}

describe('setDataValidation', () => {
	it('adds a validation to the ranges and undoes / redoes it', () => {
		const wb = withValidations();
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('A1:A10'), r('C1')], WHOLE);
		const entry = tx.commit();
		const sheet = wb.sheets[0];
		expect(sheet.dataValidations).toHaveLength(1);
		expect(sheet.dataValidations[0]).toEqual({ ...WHOLE, ranges: [r('A1:A10'), r('C1')] });
		expect(sheet.dirty).toBe(true);
		expect(findValidation(sheet, 4, 0)?.formula2).toBe('10');

		entry.undo();
		expect(wb.sheets[0].dataValidations).toEqual([]);
		entry.redo();
		expect(wb.sheets[0].dataValidations[0].type).toBe('whole');
		expect(refs(wb.sheets[0].dataValidations[0])).toEqual(['A1:A10', 'C1']);
	});

	it('replaces the validation of overwritten cells and trims the rest', () => {
		const wb = withValidations({ ...WHOLE, ranges: [r('A1:C10')] });
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('B3:B4')], LIST);
		tx.commit();
		const [kept, added] = wb.sheets[0].dataValidations;
		expect(kept.type).toBe('whole');
		expect(refs(kept)).toEqual(['A1:C2', 'A3:A4', 'C3:C4', 'A5:C10']);
		expect(added).toEqual({ ...LIST, ranges: [r('B3:B4')] });
	});

	it('drops validations whose cells are all overwritten', () => {
		const wb = withValidations({ ...WHOLE, ranges: [r('A1:A3')] }, { ...LIST, ranges: [r('D1')] });
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('A1:A5')], { ...WHOLE, formula2: '99' });
		tx.commit();
		const list = wb.sheets[0].dataValidations;
		expect(list).toHaveLength(2);
		expect(list[0]).toMatchObject({ type: 'list', ranges: [r('D1')] });
		expect(list[1]).toMatchObject({ type: 'whole', formula2: '99', ranges: [r('A1:A5')] });
	});

	it('clears validation from the ranges with a null rule (Clear All)', () => {
		const wb = withValidations(
			{ ...WHOLE, ranges: [r('A1:A10')] },
			{ ...LIST, ranges: [r('B1:B2'), r('D5')] },
			{ ...LIST, formula1: '"a,b"', ranges: [r('F1')] },
		);
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('A5:D10')], null);
		const entry = tx.commit();
		const list = wb.sheets[0].dataValidations;
		expect(list.map(refs)).toEqual([['A1:A4'], ['B1:B2'], ['F1']]);
		entry.undo();
		expect(wb.sheets[0].dataValidations.map(refs)).toEqual([['A1:A10'], ['B1:B2', 'D5'], ['F1']]);
	});

	it('keeps relative formulas pointing at the same cells when the first range is trimmed', () => {
		const wb = withValidations({ type: 'custom', formula1: 'B1>A$1', ranges: [r('A1:A10')] });
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('A1:A3')], null);
		tx.commit();
		const [dv] = wb.sheets[0].dataValidations;
		expect(refs(dv)).toEqual(['A4:A10']);
		expect(dv.formula1).toBe('B4>A$1');
	});

	it('reads relative references as written for the origin cell (the active cell)', () => {
		const wb = withValidations();
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(
			tx,
			wb,
			0,
			[r('A1:A10')],
			{ type: 'custom', formula1: 'B5>0' },
			{ origin: { row: 4, col: 0 } },
		);
		setDataValidation(tx, wb, 0, [r('C3:C4')], { ...WHOLE, formula1: 'D3', formula2: '$E$1' });
		tx.commit();
		const [custom, whole] = wb.sheets[0].dataValidations;
		expect(custom.formula1).toBe('B1>0');
		expect(whole.formula1).toBe('D3');
		expect(whole.formula2).toBe('$E$1');
	});

	it('makes overlapping input ranges disjoint and ignores a ranges field on the rule', () => {
		const wb = withValidations();
		const tx = new EditTransaction(wb, 'Data Validation');
		const rule = { ...LIST, ranges: [r('Z99')] } as DataValidationRule;
		setDataValidation(tx, wb, 0, [r('A1:B2'), r('B2:C3'), r('A1')], rule);
		tx.commit();
		const [dv] = wb.sheets[0].dataValidations;
		expect(refs(dv)).toEqual(['A1:B2', 'C2', 'B3:C3']);
	});

	it('handles whole-column selections without visiting cells', () => {
		const wb = withValidations({ ...WHOLE, ranges: [r('A1:A5'), r('B7')] });
		const tx = new EditTransaction(wb, 'Data Validation');
		setDataValidation(tx, wb, 0, [r('A:C')], LIST);
		tx.commit();
		const list = wb.sheets[0].dataValidations;
		expect(list).toHaveLength(1);
		expect(refs(list[0])).toEqual(['A:C']);
	});
});

describe('selectionValidationState', () => {
	const sheetWith = (...list: DataValidation[]) => withValidations(...list).sheets[0];

	it('reports nothing for a selection without validation', () => {
		const sheet = sheetWith({ ...WHOLE, ranges: [r('A1')] });
		expect(selectionValidationState(sheet, [r('C1:D4')])).toEqual({
			mixed: false,
			partial: false,
			covered: [],
		});
	});

	it('returns the validation covering the whole selection', () => {
		const sheet = sheetWith({ ...WHOLE, ranges: [r('A1:A10')] });
		const state = selectionValidationState(sheet, [r('A2:A4')]);
		expect(state.rule).toBe(sheet.dataValidations[0]);
		expect(state.mixed).toBe(false);
		expect(state.partial).toBe(false);
		expect(state.covered.map((x) => formatRange(x))).toEqual(['A2:A4']);
	});

	it('flags selections partly without validation', () => {
		const sheet = sheetWith({ ...WHOLE, ranges: [r('A1:A3')] });
		const state = selectionValidationState(sheet, [r('A1:A10')]);
		expect(state.partial).toBe(true);
		expect(state.mixed).toBe(false);
		expect(state.rule?.type).toBe('whole');
		expect(state.covered.map((x) => formatRange(x))).toEqual(['A1:A3']);
	});

	it('flags selections with different validations', () => {
		const sheet = sheetWith({ ...WHOLE, ranges: [r('A1:A3')] }, { ...LIST, ranges: [r('A4:A6')] });
		const state = selectionValidationState(sheet, [r('A1:A6')]);
		expect(state.mixed).toBe(true);
		expect(state.partial).toBe(false);
		expect(state.rule).toBeUndefined();
		expect(selectionValidationState(sheet, [r('A1:A10')])).toMatchObject({
			mixed: true,
			partial: true,
		});
		expect(selectionValidationState(sheet, [r('A1'), r('A5')]).mixed).toBe(true);
	});

	it('treats separate validations with the same settings as one', () => {
		const sheet = sheetWith(
			{ ...WHOLE, ranges: [r('A1:A3')] },
			{ ...WHOLE, allowBlank: true, ranges: [r('A4:A6')] },
		);
		const state = selectionValidationState(sheet, [r('A1:A6')]);
		expect(state.mixed).toBe(false);
		expect(state.partial).toBe(false);
		expect(state.rule?.ranges.map((x) => formatRange(x))).toEqual(['A1:A3', 'A4:A6']);
	});

	it('works on whole columns and multi-range selections', () => {
		const sheet = sheetWith({ ...LIST, ranges: [r('B:B')] });
		expect(selectionValidationState(sheet, [r('B:B')])).toMatchObject({
			mixed: false,
			partial: false,
		});
		expect(selectionValidationState(sheet, [r('A:C')])).toMatchObject({ partial: true });
		const multi = selectionValidationState(sheet, [r('B2'), r('B9:B10'), r('B9')]);
		expect(multi.partial).toBe(false);
		expect(multi.covered.map((x) => formatRange(x))).toEqual(['B2', 'B9:B10']);
	});
});

describe('validation settings comparison', () => {
	it('fills in defaults before comparing', () => {
		expect(
			sameValidationSettings(
				{ type: 'list', formula1: '"a"' },
				{ type: 'list', formula1: '"a"', inCellDropdown: true, allowBlank: false },
			),
		).toBe(true);
		expect(
			sameValidationSettings(
				{ type: 'whole', formula1: '1' },
				{ type: 'whole', operator: 'between', formula1: '1', errorStyle: 'stop' },
			),
		).toBe(true);
		expect(
			sameValidationSettings({ type: 'whole', formula1: '1' }, { type: 'whole', formula1: '2' }),
		).toBe(false);
		expect(sameValidationSettings({ ...LIST }, { ...LIST, prompt: 'Pick one' })).toBe(false);
	});

	it('collects every range with the same settings', () => {
		const sheet = withValidations(
			{ ...WHOLE, ranges: [r('A1:A3')] },
			{ ...LIST, ranges: [r('B1')] },
			{ ...WHOLE, ranges: [r('C1'), r('D4')] },
		).sheets[0];
		expect(sameValidationRanges(sheet, WHOLE).map((x) => formatRange(x))).toEqual([
			'A1:A3',
			'C1',
			'D4',
		]);
		expect(sameValidationRanges(sheet, { type: 'date' })).toEqual([]);
	});
});

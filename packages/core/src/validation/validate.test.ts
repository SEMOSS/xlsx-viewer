import { describe, expect, it } from 'vitest';
import { createFakeEvaluator } from '../conditional/test-utils/fake-evaluator';
import { rng, setCells } from '../conditional/test-utils/fixtures';
import { createWorkbook } from '../sheet/workbook';
import type { CellValue } from '../types/cell';
import type {
	DataValidation,
	DataValidationOperator,
	DataValidationType,
} from '../types/data-validation';
import { dateToSerial, timeToSerial } from '../values/dates';
import { ERRORS } from '../values/errors';
import { findValidation, getValidationInputMessage } from './find';
import {
	DEFAULT_VALIDATION_ERROR_MESSAGE,
	DEFAULT_VALIDATION_ERROR_TITLE,
	type DataValidationContext,
} from './types';
import { validateValue } from './validate';

function setup(cells: Record<string, CellValue> = {}) {
	const workbook = createWorkbook({ sheetNames: ['Sheet1', 'Other'] });
	setCells(workbook.sheets[0], cells);
	const evaluate = createFakeEvaluator(workbook);
	const ctx = (row = 0, col = 0): DataValidationContext => ({
		evaluate,
		sheetIndex: 0,
		row,
		col,
		date1904: false,
		workbook,
	});
	return { workbook, sheet: workbook.sheets[0], evaluate, ctx };
}

function dv(
	type: DataValidationType,
	operator?: DataValidationOperator,
	formula1?: string,
	formula2?: string,
	extra: Partial<DataValidation> = {},
): DataValidation {
	return { ranges: [rng('A1:A10')], type, operator, formula1, formula2, ...extra };
}

describe('numeric validations', () => {
	const { ctx } = setup();
	const valid = (value: CellValue, rule: DataValidation) => validateValue(value, rule, ctx()).valid;

	it.each([
		['between', '1', '10', [1, 5, 10], [0, 11]],
		['notBetween', '1', '10', [0, 11], [1, 10]],
		['equal', '5', undefined, [5], [4, 6]],
		['notEqual', '5', undefined, [4, 6], [5]],
		['greaterThan', '5', undefined, [6], [5, 4]],
		['greaterThanOrEqual', '5', undefined, [5, 6], [4]],
		['lessThan', '5', undefined, [4], [5]],
		['lessThanOrEqual', '5', undefined, [5, -3], [6]],
	] as const)('whole %s', (op, f1, f2, ok, bad) => {
		const rule = dv('whole', op, f1, f2);
		for (const v of ok) expect(valid(v, rule)).toBe(true);
		for (const v of bad) expect(valid(v, rule)).toBe(false);
	});

	it('requires whole numbers and numeric input', () => {
		const rule = dv('whole', 'between', '1', '10');
		expect(valid(2.5, rule)).toBe(false);
		expect(valid('7', rule)).toBe(true);
		expect(valid('seven', rule)).toBe(false);
		expect(valid(true, rule)).toBe(false);
		expect(valid(ERRORS.NA, rule)).toBe(false);
		expect(valid(2.5, dv('decimal', 'between', '1', '10'))).toBe(true);
		expect(valid(10.0001, dv('decimal', 'lessThanOrEqual', '10'))).toBe(false);
	});

	it('defaults to between and accepts reversed bounds', () => {
		expect(valid(5, dv('decimal', undefined, '10', '1'))).toBe(true);
		expect(valid(0.3, dv('decimal', 'equal', '0.1+0.2'))).toBe(true);
	});

	it('validates dates and times against serial bounds', () => {
		const start = dateToSerial(2024, 1, 1);
		const end = dateToSerial(2024, 12, 31);
		const dates = dv('date', 'between', String(start), String(end));
		expect(valid(dateToSerial(2024, 6, 15), dates)).toBe(true);
		expect(valid(dateToSerial(2025, 1, 1), dates)).toBe(false);
		expect(valid('6/15/2024', dates)).toBe(true);
		const times = dv(
			'time',
			'between',
			String(timeToSerial(9, 0, 0)),
			String(timeToSerial(17, 0, 0)),
		);
		expect(valid(timeToSerial(12, 30, 0), times)).toBe(true);
		expect(valid('8:59', times)).toBe(false);
		expect(valid('5:00 PM', times)).toBe(true);
	});

	it('validates text length', () => {
		const rule = dv('textLength', 'lessThanOrEqual', '3');
		expect(valid('abc', rule)).toBe(true);
		expect(valid('abcd', rule)).toBe(false);
		expect(valid(1234, rule)).toBe(false);
		expect(valid(true, dv('textLength', 'equal', '4'))).toBe(true);
		expect(valid('', dv('textLength', 'lessThan', '2'))).toBe(true);
		expect(valid(ERRORS.NA, rule)).toBe(false);
	});
});

describe('blanks and criteria formulas', () => {
	it('honours allowBlank', () => {
		const { ctx } = setup();
		expect(
			validateValue(null, dv('whole', 'between', '1', '2', { allowBlank: true }), ctx()).valid,
		).toBe(true);
		expect(
			validateValue('', dv('whole', 'between', '1', '2', { allowBlank: true }), ctx()).valid,
		).toBe(true);
		expect(validateValue(null, dv('whole', 'between', '1', '2'), ctx()).valid).toBe(false);
		expect(validateValue(null, dv('list', undefined, '"a,b"'), ctx()).valid).toBe(false);
	});

	it('evaluates criteria relative to the validated cell', () => {
		const { ctx } = setup({ B1: 10, B2: 20 });
		const rule = dv('whole', 'lessThanOrEqual', 'B1');
		expect(validateValue(15, rule, ctx(0, 0)).valid).toBe(false);
		expect(validateValue(15, rule, ctx(1, 0)).valid).toBe(true);
		expect(validateValue(15, dv('whole', 'lessThanOrEqual', '$B$1'), ctx(1, 0)).valid).toBe(false);
	});

	it('accepts anything when a criterion cell is blank and blanks are ignored', () => {
		const { ctx } = setup({ B1: 10 });
		expect(
			validateValue(500, dv('whole', 'between', '$B$1', '$C$1', { allowBlank: true }), ctx()).valid,
		).toBe(true);
		expect(validateValue(500, dv('whole', 'between', '$B$1', '$C$1'), ctx()).valid).toBe(false);
		expect(validateValue(5, dv('whole', 'between', '$B$1', '$C$1'), ctx()).valid).toBe(true);
	});

	it('rejects values when a criterion is an error and imposes nothing when missing', () => {
		const { ctx } = setup();
		expect(validateValue(1, dv('whole', 'greaterThan', 'NA()'), ctx()).valid).toBe(false);
		expect(validateValue(1, dv('whole', 'greaterThan', '"abc"'), ctx()).valid).toBe(false);
		expect(validateValue(1, dv('whole', 'greaterThan'), ctx()).valid).toBe(true);
		expect(validateValue('anything', dv('none'), ctx()).valid).toBe(true);
	});
});

describe('custom validations', () => {
	it('evaluates the formula with the pending value in the cell', () => {
		const { ctx, sheet } = setup({ A1: 'old', A2: 'x' });
		const rule = dv('custom', undefined, 'ISNUMBER(A1)');
		expect(validateValue(5, rule, ctx(0, 0)).valid).toBe(true);
		expect(validateValue('text', rule, ctx(0, 0)).valid).toBe(false);
		expect(sheet.cells.get(0)?.value).toBe('old');
		const unique = dv('custom', undefined, 'COUNTIF($A$1:$A$10,A1)=1');
		expect(validateValue('x', unique, ctx(2, 0)).valid).toBe(false);
		expect(validateValue('y', unique, ctx(2, 0)).valid).toBe(true);
		expect(sheet.cells.has(2 * 16384)).toBe(false);
	});

	it('treats numbers as conditions, errors as invalid, and works without a workbook', () => {
		const { ctx, evaluate } = setup();
		expect(validateValue(1, dv('custom', undefined, '2'), ctx()).valid).toBe(true);
		expect(validateValue(1, dv('custom', undefined, 'NA()'), ctx()).valid).toBe(false);
		expect(validateValue(1, dv('custom', undefined, 'THROW()'), ctx()).valid).toBe(false);
		const bare: DataValidationContext = {
			evaluate,
			sheetIndex: 0,
			row: 0,
			col: 0,
			date1904: false,
		};
		expect(validateValue(1, dv('custom', undefined, 'ROW()=1'), bare).valid).toBe(true);
	});
});

describe('alerts', () => {
	const { ctx } = setup();
	const alert = { showErrorMessage: true };
	it('uses Excel defaults', () => {
		expect(validateValue(0, dv('whole', 'greaterThan', '1', undefined, alert), ctx())).toEqual({
			valid: false,
			style: 'stop',
			title: DEFAULT_VALIDATION_ERROR_TITLE,
			message: DEFAULT_VALIDATION_ERROR_MESSAGE,
		});
		expect(DEFAULT_VALIDATION_ERROR_MESSAGE).toBe(
			"This value doesn't match the data validation restrictions defined for this cell.",
		);
		expect(validateValue(5, dv('whole', 'greaterThan', '1'), ctx())).toEqual({ valid: true });
	});

	it('uses the rule title, message and style', () => {
		const rule = dv('whole', 'greaterThan', '1', undefined, {
			...alert,
			errorStyle: 'warning',
			errorTitle: 'Too small',
			error: 'Enter more than 1',
		});
		expect(validateValue(0, rule, ctx())).toEqual({
			valid: false,
			style: 'warning',
			title: 'Too small',
			message: 'Enter more than 1',
		});
		expect(validateValue(0, { ...rule, errorStyle: 'information' }, ctx()).style).toBe(
			'information',
		);
	});

	it('shows no alert unless the rule enables it, like Excel', () => {
		expect(
			validateValue(
				0,
				dv('whole', 'greaterThan', '1', undefined, { showErrorMessage: false }),
				ctx(),
			),
		).toEqual({ valid: false });
		expect(validateValue(0, dv('whole', 'greaterThan', '1'), ctx())).toEqual({ valid: false });
	});
});

describe('findValidation / getInputMessage', () => {
	it('finds the rule covering a cell', () => {
		const { sheet } = setup();
		const a = dv('whole', 'greaterThan', '1');
		const b: DataValidation = { ranges: [rng('C1'), rng('E5:F6')], type: 'list', formula1: '"x"' };
		sheet.dataValidations.push(a, b);
		expect(findValidation(sheet, 3, 0)).toBe(a);
		expect(findValidation(sheet, 5, 5)).toBe(b);
		expect(findValidation(sheet, 0, 2)).toBe(b);
		expect(findValidation(sheet, 0, 1)).toBeUndefined();
	});

	it('returns the input message', () => {
		const show = { showInputMessage: true };
		expect(
			getValidationInputMessage(
				dv('none', undefined, undefined, undefined, {
					...show,
					promptTitle: 'Hint',
					prompt: 'Type a number',
				}),
			),
		).toEqual({ title: 'Hint', message: 'Type a number' });
		expect(
			getValidationInputMessage(
				dv('none', undefined, undefined, undefined, { ...show, prompt: 'p' }),
			),
		).toEqual({ message: 'p' });
		expect(
			getValidationInputMessage(dv('none', undefined, undefined, undefined, { prompt: 'p' })),
		).toBeUndefined();
		expect(
			getValidationInputMessage(
				dv('none', undefined, undefined, undefined, { prompt: 'p', showInputMessage: false }),
			),
		).toBeUndefined();
		expect(
			getValidationInputMessage(dv('none', undefined, undefined, undefined, show)),
		).toBeUndefined();
	});
});

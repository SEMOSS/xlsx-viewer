import { describe, expect, it } from 'vitest';

import { dateToSerial } from '../values/dates';
import {
	checkValidationCriteria,
	validationCriteriaToFormula,
	validationFormulaToText,
	validationOperatorTakesTwo,
	validationUsesOperator,
	validationValueLabels,
	VALIDATION_LIST_SOURCE_MESSAGE,
} from './index';

describe('validationValueLabels', () => {
	it('follows the labels of Excel for each type and operator', () => {
		expect(validationValueLabels('none')).toEqual([]);
		expect(validationValueLabels('whole', 'between')).toEqual(['Minimum', 'Maximum']);
		expect(validationValueLabels('decimal', 'notBetween')).toEqual(['Minimum', 'Maximum']);
		expect(validationValueLabels('whole', 'equal')).toEqual(['Value']);
		expect(validationValueLabels('decimal', 'greaterThan')).toEqual(['Minimum']);
		expect(validationValueLabels('decimal', 'lessThanOrEqual')).toEqual(['Maximum']);
		expect(validationValueLabels('textLength', 'notEqual')).toEqual(['Length']);
		expect(validationValueLabels('date', 'between')).toEqual(['Start date', 'End date']);
		expect(validationValueLabels('date', 'greaterThanOrEqual')).toEqual(['Start date']);
		expect(validationValueLabels('date', 'lessThan')).toEqual(['End date']);
		expect(validationValueLabels('date', 'equal')).toEqual(['Date']);
		expect(validationValueLabels('time', 'between')).toEqual(['Start time', 'End time']);
		expect(validationValueLabels('time', 'equal')).toEqual(['Time']);
		expect(validationValueLabels('list', 'greaterThan')).toEqual(['Source']);
		expect(validationValueLabels('custom')).toEqual(['Formula']);
	});

	it('tells which types take an operator and which operators take two values', () => {
		expect(
			['whole', 'decimal', 'date', 'time', 'textLength'].every((t) =>
				validationUsesOperator(t as 'whole'),
			),
		).toBe(true);
		expect(validationUsesOperator('list')).toBe(false);
		expect(validationUsesOperator('custom')).toBe(false);
		expect(validationUsesOperator('none')).toBe(false);
		expect(validationOperatorTakesTwo('between')).toBe(true);
		expect(validationOperatorTakesTwo('notBetween')).toBe(true);
		expect(validationOperatorTakesTwo('equal')).toBe(false);
	});
});

describe('validationCriteriaToFormula', () => {
	it('turns references and formulas into display-form formulas', () => {
		expect(validationCriteriaToFormula('whole', '=A1')).toBe('A1');
		expect(validationCriteriaToFormula('whole', ' = $B$2 ')).toBe('$B$2');
		expect(validationCriteriaToFormula('list', '=$A$1:$A$5')).toBe('$A$1:$A$5');
		expect(validationCriteriaToFormula('list', '=Sheet2!$A$1:$A$5')).toBe('Sheet2!$A$1:$A$5');
		expect(validationCriteriaToFormula('custom', '=ISNUMBER(A1)')).toBe('ISNUMBER(A1)');
		expect(validationCriteriaToFormula('custom', 'A1>0')).toBe('A1>0');
	});

	it('returns undefined for empty boxes and Any value', () => {
		expect(validationCriteriaToFormula('whole', '')).toBeUndefined();
		expect(validationCriteriaToFormula('whole', '   ')).toBeUndefined();
		expect(validationCriteriaToFormula('whole', '=')).toBeUndefined();
		expect(validationCriteriaToFormula('none', '5')).toBeUndefined();
	});

	it('normalises numbers and keeps unparseable text', () => {
		expect(validationCriteriaToFormula('whole', '10')).toBe('10');
		expect(validationCriteriaToFormula('decimal', '1,000.5')).toBe('1000.5');
		expect(validationCriteriaToFormula('decimal', '50%')).toBe('0.5');
		expect(validationCriteriaToFormula('textLength', ' 3 ')).toBe('3');
		expect(validationCriteriaToFormula('whole', 'abc')).toBe('abc');
	});

	it('stores dates as serial numbers and times as day fractions', () => {
		expect(validationCriteriaToFormula('date', '1/15/2024')).toBe(
			String(dateToSerial(2024, 1, 15)),
		);
		expect(validationCriteriaToFormula('date', '2024-01-15')).toBe('45306');
		expect(validationCriteriaToFormula('date', '1/15/2024', true)).toBe(
			String(dateToSerial(2024, 1, 15, true)),
		);
		expect(validationCriteriaToFormula('date', '45306')).toBe('45306');
		expect(validationCriteriaToFormula('time', '12:00')).toBe('0.5');
		expect(validationCriteriaToFormula('time', '9:30')).toBe('0.395833333333333');
		expect(validationCriteriaToFormula('time', '9:30 PM')).toBe('0.895833333333333');
		expect(validationCriteriaToFormula('time', '0.25')).toBe('0.25');
	});

	it('quotes typed lists and escapes quotes', () => {
		expect(validationCriteriaToFormula('list', 'Yes,No')).toBe('"Yes,No"');
		expect(validationCriteriaToFormula('list', ' Low, High ')).toBe('"Low, High"');
		expect(validationCriteriaToFormula('list', 'say "hi",bye')).toBe('"say ""hi"",bye"');
	});
});

describe('validationFormulaToText', () => {
	it('shows serials as dates and fractions as times', () => {
		expect(validationFormulaToText('date', '45306')).toBe('1/15/2024');
		expect(validationFormulaToText('date', '45306.5')).toBe('1/15/2024 12:00 PM');
		expect(validationFormulaToText('date', String(dateToSerial(2024, 1, 15, true)), true)).toBe(
			'1/15/2024',
		);
		expect(validationFormulaToText('time', '0.395833333333333')).toBe('9:30 AM');
		expect(validationFormulaToText('time', '0.895833333333333')).toBe('9:30 PM');
		expect(validationFormulaToText('time', '0')).toBe('12:00 AM');
		expect(validationFormulaToText('time', '0.5')).toBe('12:00 PM');
		expect(validationFormulaToText('time', String(15 / 86400 + 0.5))).toBe('12:00:15 PM');
	});

	it('shows literal lists without quotes and references with =', () => {
		expect(validationFormulaToText('list', '"Yes,No"')).toBe('Yes,No');
		expect(validationFormulaToText('list', '"say ""hi"",bye"')).toBe('say "hi",bye');
		expect(validationFormulaToText('list', '$A$1:$A$5')).toBe('=$A$1:$A$5');
		expect(validationFormulaToText('list', 'MyList')).toBe('=MyList');
		expect(validationFormulaToText('custom', 'ISNUMBER(A1)')).toBe('=ISNUMBER(A1)');
		expect(validationFormulaToText('whole', 'A1')).toBe('=A1');
		expect(validationFormulaToText('date', 'TODAY()')).toBe('=TODAY()');
	});

	it('shows numbers as numbers and empty formulas as empty text', () => {
		expect(validationFormulaToText('whole', '10')).toBe('10');
		expect(validationFormulaToText('decimal', '-2.5')).toBe('-2.5');
		expect(validationFormulaToText('textLength', '255')).toBe('255');
		expect(validationFormulaToText('whole', undefined)).toBe('');
		expect(validationFormulaToText('whole', '  ')).toBe('');
	});

	it('round-trips with validationCriteriaToFormula', () => {
		const cases: Array<['date' | 'time' | 'list' | 'whole' | 'custom', string]> = [
			['date', '45306'],
			['date', '45306.75'],
			['time', '0.395833333333333'],
			['time', '0.75'],
			['list', '"Yes,No,Maybe"'],
			['list', '$A$1:$A$5'],
			['whole', '42'],
			['custom', 'AND(A1>0,A1<10)'],
		];
		for (const [type, formula] of cases) {
			expect(validationCriteriaToFormula(type, validationFormulaToText(type, formula))).toBe(
				formula,
			);
		}
	});
});

describe('checkValidationCriteria', () => {
	it('accepts Any value and complete criteria', () => {
		expect(checkValidationCriteria({ type: 'none', value1: '' })).toBeNull();
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '1', value2: '10' }),
		).toBeNull();
		expect(
			checkValidationCriteria({ type: 'decimal', operator: 'greaterThan', value1: '0.5' }),
		).toBeNull();
		expect(
			checkValidationCriteria({
				type: 'date',
				operator: 'between',
				value1: '1/1/2024',
				value2: '12/31/2024',
			}),
		).toBeNull();
		expect(
			checkValidationCriteria({ type: 'time', operator: 'lessThan', value1: '5:00 PM' }),
		).toBeNull();
		expect(
			checkValidationCriteria({ type: 'textLength', operator: 'lessThanOrEqual', value1: '10' }),
		).toBeNull();
		expect(checkValidationCriteria({ type: 'list', value1: 'Yes,No' })).toBeNull();
		expect(checkValidationCriteria({ type: 'list', value1: '=$A$1:$A$5' })).toBeNull();
		expect(checkValidationCriteria({ type: 'list', value1: '=Sheet2!$A$1:$E$1' })).toBeNull();
		expect(checkValidationCriteria({ type: 'list', value1: '=MyList' })).toBeNull();
		expect(checkValidationCriteria({ type: 'custom', value1: '=A1>0' })).toBeNull();
	});

	it('asks for missing values using the box label', () => {
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '', value2: '5' }),
		).toBe('You must enter a value in the Minimum box.');
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '1', value2: ' ' }),
		).toBe('You must enter a value in the Maximum box.');
		expect(checkValidationCriteria({ type: 'date', operator: 'equal', value1: '' })).toBe(
			'You must enter a value in the Date box.',
		);
		expect(checkValidationCriteria({ type: 'custom', value1: '=' })).toBe(
			'You must enter a value in the Formula box.',
		);
	});

	it('rejects values that do not fit the type', () => {
		expect(checkValidationCriteria({ type: 'whole', operator: 'greaterThan', value1: 'abc' })).toBe(
			'The Minimum you entered is invalid.',
		);
		expect(checkValidationCriteria({ type: 'whole', operator: 'equal', value1: '1.5' })).toBe(
			'The Value you entered is invalid.',
		);
		expect(
			checkValidationCriteria({ type: 'textLength', operator: 'lessThan', value1: '-1' }),
		).toBe('The Maximum you entered is invalid.');
		expect(
			checkValidationCriteria({
				type: 'date',
				operator: 'between',
				value1: 'soon',
				value2: '1/1/2024',
			}),
		).toBe('The Start date you entered is invalid.');
		expect(
			checkValidationCriteria({
				type: 'time',
				operator: 'between',
				value1: '9:00',
				value2: 'later',
			}),
		).toBe('The End time you entered is invalid.');
	});

	it('requires the upper bound to be at least the lower bound', () => {
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '10', value2: '1' }),
		).toBe('The Maximum must be greater than or equal to the Minimum.');
		expect(
			checkValidationCriteria({
				type: 'date',
				operator: 'notBetween',
				value1: '2/1/2024',
				value2: '1/1/2024',
			}),
		).toBe('The End date must be greater than or equal to the Start date.');
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '5', value2: '5' }),
		).toBeNull();
		// Formulas are not range-checked.
		expect(
			checkValidationCriteria({ type: 'whole', operator: 'between', value1: '=B1', value2: '1' }),
		).toBeNull();
	});

	it('validates list sources', () => {
		expect(checkValidationCriteria({ type: 'list', value1: '' })).toBe(
			VALIDATION_LIST_SOURCE_MESSAGE,
		);
		expect(checkValidationCriteria({ type: 'list', value1: '=' })).toBe(
			VALIDATION_LIST_SOURCE_MESSAGE,
		);
		expect(checkValidationCriteria({ type: 'list', value1: '=$A$1:$B$5' })).toBe(
			VALIDATION_LIST_SOURCE_MESSAGE,
		);
		expect(checkValidationCriteria({ type: 'list', value1: "='My Sheet'!A1:C3" })).toBe(
			VALIDATION_LIST_SOURCE_MESSAGE,
		);
		expect(VALIDATION_LIST_SOURCE_MESSAGE).toBe(
			'The list source must be a delimited list, or a reference to single row or column.',
		);
	});
});

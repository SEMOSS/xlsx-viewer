import { parseRange } from '../address/range';
import { numberToText } from '../formula/coerce';
import type { DataValidationOperator, DataValidationType } from '../types/data-validation';
import { parseLiteralList } from '../validation/list-source';
import { MAX_DATE_SERIAL, serialToDateParts } from '../values/dates';
import { parseNumericText } from '../values/parse-number';

/**
 * Conversions between the text typed in Excel's Data Validation dialog and
 * the criteria formulas stored on a rule, plus the checks Excel runs when
 * the dialog is confirmed.
 */

const OPERATOR_TYPES: ReadonlySet<DataValidationType> = new Set([
	'whole',
	'decimal',
	'date',
	'time',
	'textLength',
]);

/** Excel's message for a missing or two-dimensional list source. */
export const VALIDATION_LIST_SOURCE_MESSAGE =
	'The list source must be a delimited list, or a reference to single row or column.';

/** The rule type takes an operator (the dialog's "Data" box). */
export function validationUsesOperator(type: DataValidationType): boolean {
	return OPERATOR_TYPES.has(type);
}

/** The operator compares against two values (Minimum and Maximum). */
export function validationOperatorTakesTwo(operator: DataValidationOperator): boolean {
	return operator === 'between' || operator === 'notBetween';
}

const BOUND_LABELS: Record<string, [string, string, string]> = {
	whole: ['Minimum', 'Maximum', 'Value'],
	decimal: ['Minimum', 'Maximum', 'Value'],
	textLength: ['Minimum', 'Maximum', 'Length'],
	date: ['Start date', 'End date', 'Date'],
	time: ['Start time', 'End time', 'Time'],
};

/**
 * Labels of the dialog's value boxes for a rule type and operator, in
 * formula order (`formula1`, then `formula2`): `['Minimum', 'Maximum']` for
 * whole numbers between two values, `['Start date']` for dates greater than
 * a value, `['Source']` for lists, `['Formula']` for custom rules and none
 * for "Any value".
 */
export function validationValueLabels(
	type: DataValidationType,
	operator: DataValidationOperator = 'between',
): string[] {
	if (type === 'list') {
		return ['Source'];
	}
	if (type === 'custom') {
		return ['Formula'];
	}
	const labels = BOUND_LABELS[type];
	if (!labels) {
		return [];
	}
	const [min, max, single] = labels;
	switch (operator) {
		case 'between':
		case 'notBetween':
			return [min, max];
		case 'equal':
		case 'notEqual':
			return [single];
		case 'greaterThan':
		case 'greaterThanOrEqual':
			return [min];
		default:
			return [max];
	}
}

/**
 * Stored criteria formula (display form, no `=`) for the text typed in a
 * value box, or `undefined` when the box is empty. `=A1` and `=$A$1:$A$5`
 * become references; dates (`1/15/2024`) and times (`9:30 PM`) become their
 * serial numbers; a typed list (`Yes,No`) becomes a quoted literal
 * (`"Yes,No"`); numbers are normalised (`1,000` is `1000`).
 */
export function validationCriteriaToFormula(
	type: DataValidationType,
	text: string,
	date1904 = false,
): string | undefined {
	const t = text.trim();
	if (t === '' || type === 'none') {
		return undefined;
	}
	if (t.startsWith('=')) {
		const formula = t.slice(1).trim();
		return formula === '' ? undefined : formula;
	}
	if (type === 'list') {
		return `"${t.replace(/"/g, '""')}"`;
	}
	if (type === 'custom') {
		return t;
	}
	const parsed = parseNumericText(t, { date1904 });
	return parsed ? numberToText(parsed.value) : t;
}

const PLAIN_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

function dateText(serial: number, date1904: boolean): string {
	const p = serialToDateParts(serial, date1904);
	const date = `${p.month}/${p.day}/${p.year}`;
	return serial === Math.floor(serial) ? date : `${date} ${timeOfDayText(serial, date1904)}`;
}

function timeOfDayText(serial: number, date1904: boolean): string {
	const p = serialToDateParts(serial, date1904);
	const hour12 = p.hour % 12 === 0 ? 12 : p.hour % 12;
	const minutes = String(p.minute).padStart(2, '0');
	const seconds = p.second ? `:${String(p.second).padStart(2, '0')}` : '';
	return `${hour12}:${minutes}${seconds} ${p.hour < 12 ? 'AM' : 'PM'}`;
}

/**
 * Text to show in a value box for a stored criteria formula (the inverse of
 * {@link validationCriteriaToFormula}): date serials as `m/d/yyyy`, time
 * fractions as `h:mm AM/PM`, literal lists as `Yes,No`, numbers as numbers
 * and references or formulas with a leading `=`.
 */
export function validationFormulaToText(
	type: DataValidationType,
	formula: string | undefined,
	date1904 = false,
): string {
	const f = formula?.trim() ?? '';
	if (f === '') {
		return '';
	}
	if (type === 'list') {
		const items = parseLiteralList(f);
		return items ? items.join(',') : `=${f}`;
	}
	if (type === 'custom' || type === 'none' || !PLAIN_NUMBER.test(f)) {
		return `=${f}`;
	}
	const n = Number(f);
	if (!Number.isFinite(n)) {
		return `=${f}`;
	}
	if ((type === 'date' || type === 'time') && n >= 0 && n <= MAX_DATE_SERIAL + 1) {
		if (type === 'date') {
			return dateText(n, date1904);
		}
		return n < 1 ? timeOfDayText(n, date1904) : dateText(n, date1904);
	}
	return numberToText(n);
}

/** Values typed in the Data Validation dialog, as checked on OK. */
export interface ValidationCriteriaInput {
	type: DataValidationType;
	operator?: DataValidationOperator;
	/** Text of the first value box (Minimum, Value, Start date, Source, Formula, ...). */
	value1: string;
	/** Text of the second value box (Maximum, End date, End time). */
	value2?: string;
	date1904?: boolean;
}

/** Number for a constant typed in a value box, or `undefined` when it does not fit the type. */
function criterionNumber(
	type: DataValidationType,
	text: string,
	date1904: boolean,
): number | undefined {
	const parsed = parseNumericText(text, { date1904 });
	if (!parsed || !Number.isFinite(parsed.value)) {
		return undefined;
	}
	const n = parsed.value;
	switch (type) {
		case 'whole':
			return Number.isInteger(n) ? n : undefined;
		case 'textLength':
			return Number.isInteger(n) && n >= 0 ? n : undefined;
		case 'date':
			return n >= 0 && n <= MAX_DATE_SERIAL + 1 ? n : undefined;
		case 'time':
			return n >= 0 ? n : undefined;
		default:
			return n;
	}
}

/** A list source is a literal list, a single row or column reference, or a formula / name. */
function listSourceError(text: string): string | null {
	if (text === '' || text === '=') {
		return VALIDATION_LIST_SOURCE_MESSAGE;
	}
	if (!text.startsWith('=')) {
		return null;
	}
	const ref = text
		.slice(1)
		.trim()
		.replace(/^(?:'(?:[^']|'')+'|[^'!"(),]+)!/, '');
	const range = parseRange(ref);
	if (range && range.bottom > range.top && range.right > range.left) {
		return VALIDATION_LIST_SOURCE_MESSAGE;
	}
	return null;
}

/**
 * Check the dialog's criteria like Excel does on OK. Returns the message
 * Excel shows (`You must enter a value in the Minimum box.`, `The Start
 * date you entered is invalid.`, `The Maximum must be greater than or equal
 * to the Minimum.`, or the list source message), or `null` when the values
 * are acceptable. Values starting with `=` are formulas and are not
 * range-checked.
 */
export function checkValidationCriteria(input: ValidationCriteriaInput): string | null {
	const { type } = input;
	const first = input.value1.trim();
	if (type === 'none') {
		return null;
	}
	if (type === 'list') {
		return listSourceError(first);
	}
	const labels = validationValueLabels(type, input.operator);
	const texts = [first, (input.value2 ?? '').trim()];
	const numbers: Array<number | undefined> = [];
	for (let i = 0; i < labels.length; i++) {
		const text = texts[i];
		if (text === '' || text === '=') {
			return `You must enter a value in the ${labels[i]} box.`;
		}
		if (type === 'custom' || text.startsWith('=')) {
			numbers.push(undefined);
			continue;
		}
		const n = criterionNumber(type, text, input.date1904 === true);
		if (n === undefined) {
			return `The ${labels[i]} you entered is invalid.`;
		}
		numbers.push(n);
	}
	const [low, high] = numbers;
	if (labels.length === 2 && low !== undefined && high !== undefined && high < low) {
		return `The ${labels[1]} must be greater than or equal to the ${labels[0]}.`;
	}
	return null;
}

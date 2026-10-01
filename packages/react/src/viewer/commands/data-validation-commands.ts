import type {
	CellAddress,
	DataValidation,
	DataValidationOperator,
	DataValidationRule,
	DataValidationType,
	RangeAddress,
	SelectionValidationState,
} from 'xlsx-viewer-core';
import {
	checkValidationCriteria,
	sameValidationRanges,
	selectionValidationState,
	setDataValidation,
	translateFormula,
	validationCriteriaToFormula,
	validationFormulaToText,
	validationUsesOperator,
	validationValueLabels,
} from 'xlsx-viewer-core';

import type { Translate, TranslationKey } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';

/** Error alert style of a validation. */
export type ValidationErrorStyle = NonNullable<DataValidation['errorStyle']>;

/** The fields of the Data Validation dialog (all three tabs). */
export interface ValidationForm {
	type: DataValidationType;
	operator: DataValidationOperator;
	/** First value box (Minimum, Value, Source, Formula, ...), as typed. */
	value1: string;
	/** Second value box (Maximum, End date, End time), as typed. */
	value2: string;
	ignoreBlank: boolean;
	inCellDropdown: boolean;
	showInputMessage: boolean;
	promptTitle: string;
	prompt: string;
	showErrorMessage: boolean;
	errorStyle: ValidationErrorStyle;
	errorTitle: string;
	error: string;
}

/** Excel's defaults for a new validation (also what Clear All restores). */
export const DEFAULT_VALIDATION_FORM: ValidationForm = {
	type: 'none',
	operator: 'between',
	value1: '',
	value2: '',
	ignoreBlank: true,
	inCellDropdown: true,
	showInputMessage: true,
	promptTitle: '',
	prompt: '',
	showErrorMessage: true,
	errorStyle: 'stop',
	errorTitle: '',
	error: '',
};

/** Excel limits titles to 32 and messages to 255 characters. */
export const VALIDATION_TITLE_MAX = 32;
export const VALIDATION_MESSAGE_MAX = 255;

/** Caption keys of the value boxes, by the label core's `validationValueLabels` gives them. */
const VALUE_BOX_KEYS: ReadonlyMap<string, TranslationKey> = new Map([
	['Minimum', 'dialog.dataValidation.box.minimum'],
	['Maximum', 'dialog.dataValidation.box.maximum'],
	['Value', 'dialog.dataValidation.box.value'],
	['Length', 'dialog.dataValidation.box.length'],
	['Start date', 'dialog.dataValidation.box.startDate'],
	['End date', 'dialog.dataValidation.box.endDate'],
	['Date', 'dialog.dataValidation.box.date'],
	['Start time', 'dialog.dataValidation.box.startTime'],
	['End time', 'dialog.dataValidation.box.endTime'],
	['Time', 'dialog.dataValidation.box.time'],
	['Source', 'dialog.dataValidation.box.source'],
	['Formula', 'dialog.dataValidation.box.formula'],
]);

/**
 * Captions of the dialog's value boxes for a rule type and operator, in
 * formula order: `['Minimum:', 'Maximum:']` for whole numbers between two
 * values, `['Source:']` for lists, none for "Any value".
 */
export function validationValueCaptions(
	type: DataValidationType,
	operator: DataValidationOperator,
	t: Translate,
): string[] {
	return validationValueLabels(type, operator).map((label) => {
		const key = VALUE_BOX_KEYS.get(label);
		return key ? t(key) : `${label}:`;
	});
}

/**
 * Dialog fields for an existing validation. Relative references are shown
 * as seen from `at` (the active cell), like Excel.
 */
export function validationFormFromRule(
	rule: DataValidation,
	at: CellAddress,
	date1904: boolean,
): ValidationForm {
	const anchor = rule.ranges[0];
	const dr = anchor ? at.row - anchor.top : 0;
	const dc = anchor ? at.col - anchor.left : 0;
	const text = (formula: string | undefined) =>
		validationFormulaToText(
			rule.type,
			formula === undefined ? undefined : translateFormula(formula, dr, dc),
			date1904,
		);
	return {
		type: rule.type,
		operator: rule.operator ?? 'between',
		value1: text(rule.formula1),
		value2: text(rule.formula2),
		ignoreBlank: rule.allowBlank === true,
		inCellDropdown: rule.inCellDropdown !== false,
		showInputMessage: rule.showInputMessage === true,
		promptTitle: rule.promptTitle ?? '',
		prompt: rule.prompt ?? '',
		showErrorMessage: rule.showErrorMessage === true,
		errorStyle: rule.errorStyle ?? 'stop',
		errorTitle: rule.errorTitle ?? '',
		error: rule.error ?? '',
	};
}

/** A formula refers to another sheet (Excel then stores the rule in the 2010 extension). */
function refersToOtherSheet(formula: string | undefined): boolean {
	return formula !== undefined && formula.replace(/"(?:[^"]|"")*"/g, '').includes('!');
}

function nonBlank(value: string): string | undefined {
	return value.trim() === '' ? undefined : value;
}

/**
 * The rule to store for the dialog fields, or `null` to remove validation:
 * "Any value" without an input message is no validation at all.
 */
export function ruleFromValidationForm(
	form: ValidationForm,
	date1904: boolean,
): DataValidationRule | null {
	const promptTitle = nonBlank(form.promptTitle);
	const prompt = nonBlank(form.prompt);
	if (form.type === 'none' && !promptTitle && !prompt) {
		return null;
	}
	const labels = validationValueLabels(form.type, form.operator);
	const formula1 =
		labels.length > 0 ? validationCriteriaToFormula(form.type, form.value1, date1904) : undefined;
	const formula2 =
		labels.length > 1 ? validationCriteriaToFormula(form.type, form.value2, date1904) : undefined;
	const rule: DataValidationRule = {
		type: form.type,
		allowBlank: form.ignoreBlank,
		showInputMessage: form.showInputMessage,
		showErrorMessage: form.showErrorMessage,
	};
	if (validationUsesOperator(form.type)) {
		rule.operator = form.operator;
	}
	if (formula1 !== undefined) {
		rule.formula1 = formula1;
	}
	if (formula2 !== undefined) {
		rule.formula2 = formula2;
	}
	if (form.type === 'list') {
		rule.inCellDropdown = form.inCellDropdown;
	}
	rule.errorStyle = form.errorStyle;
	const extra = {
		errorTitle: nonBlank(form.errorTitle),
		error: nonBlank(form.error),
		promptTitle,
		prompt,
	};
	for (const [key, value] of Object.entries(extra) as Array<
		[keyof typeof extra, string | undefined]
	>) {
		if (value !== undefined) {
			rule[key] = value;
		}
	}
	if (refersToOtherSheet(formula1) || refersToOtherSheet(formula2)) {
		rule.fromExtension = true;
	}
	return rule;
}

/** What the dialog starts from: the selection, what it contains and the prefilled fields. */
export interface DataValidationSetup {
	selection: RangeAddress[];
	state: SelectionValidationState;
	/** Fields of the selection's validation (defaults when there is none or it is mixed). */
	form: ValidationForm;
}

export function dataValidationSetup(store: ViewerStore): DataValidationSetup {
	const sel = store.selectionOf();
	const state = selectionValidationState(store.activeSheet, sel.ranges);
	const form =
		state.rule && !state.mixed
			? validationFormFromRule(state.rule, sel.active, store.workbook.date1904)
			: DEFAULT_VALIDATION_FORM;
	return { selection: sel.ranges, state, form };
}

/**
 * OK in the Data Validation dialog: check the criteria (returns Excel's
 * message when they are not acceptable), then set or remove the validation
 * of `targets` in one undoable step. With `sameAs`, every other cell with
 * the same settings as that rule is changed too.
 */
export function applyValidationForm(
	store: ViewerStore,
	form: ValidationForm,
	targets: readonly RangeAddress[],
	sameAs?: DataValidation,
): string | null {
	const date1904 = store.workbook.date1904;
	const error = checkValidationCriteria({
		type: form.type,
		operator: form.operator,
		value1: form.value1,
		value2: form.value2,
		date1904,
	});
	if (error) {
		return error;
	}
	const rule = ruleFromValidationForm(form, date1904);
	const sheetIndex = store.state.activeSheet;
	const ranges = sameAs
		? [...targets, ...sameValidationRanges(store.activeSheet, sameAs)]
		: [...targets];
	const origin = store.selectionOf().active;
	store.transact(store.t('cmd.dataValidation'), (tx) => {
		setDataValidation(tx, store.workbook, sheetIndex, ranges, rule, { origin });
	});
	return null;
}

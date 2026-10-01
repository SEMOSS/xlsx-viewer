import { parseSqref } from '../address/range';
import { stripFormulaPrefixes } from '../formula/prefixes';
import type {
	DataValidation,
	DataValidationOperator,
	DataValidationType,
} from '../types/data-validation';
import { textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { bool, kid, kids, oneOf, str } from './xml-utils';

const TYPES: readonly DataValidationType[] = [
	'none',
	'whole',
	'decimal',
	'list',
	'date',
	'time',
	'textLength',
	'custom',
];

const OPERATORS: readonly DataValidationOperator[] = [
	'between',
	'notBetween',
	'equal',
	'notEqual',
	'lessThan',
	'lessThanOrEqual',
	'greaterThan',
	'greaterThanOrEqual',
];

/** Formula of `formula1` / `formula2`: its text (2007) or its `xm:f` child (2010). */
function formulaText(el: XmlElement | undefined): string | undefined {
	if (!el) {
		return undefined;
	}
	const f = kid(el, 'f');
	return stripFormulaPrefixes(textContent(f ?? el));
}

function parseValidation(
	el: XmlElement,
	sqref: string,
	fromExtension: boolean,
): DataValidation | undefined {
	const ranges = parseSqref(sqref);
	if (ranges.length === 0) {
		return undefined;
	}
	const type = oneOf(str(el, 'type'), TYPES) ?? 'none';
	const dv: DataValidation = { ranges, type };
	const operator = oneOf(str(el, 'operator'), OPERATORS);
	if (operator) {
		dv.operator = operator;
	}
	const f1 = formulaText(kid(el, 'formula1'));
	if (f1 !== undefined) {
		dv.formula1 = f1;
	}
	const f2 = formulaText(kid(el, 'formula2'));
	if (f2 !== undefined) {
		dv.formula2 = f2;
	}
	const allowBlank = bool(el, 'allowBlank');
	if (allowBlank !== undefined) {
		dv.allowBlank = allowBlank;
	}
	// OOXML `showDropDown="1"` HIDES the in-cell dropdown.
	const suppress = bool(el, 'showDropDown');
	if (type === 'list' || suppress !== undefined) {
		dv.inCellDropdown = !(suppress ?? false);
	}
	const showInput = bool(el, 'showInputMessage');
	if (showInput !== undefined) {
		dv.showInputMessage = showInput;
	}
	const showError = bool(el, 'showErrorMessage');
	if (showError !== undefined) {
		dv.showErrorMessage = showError;
	}
	const errorStyle = oneOf(str(el, 'errorStyle'), ['stop', 'warning', 'information'] as const);
	if (errorStyle) {
		dv.errorStyle = errorStyle;
	}
	for (const key of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
		const v = str(el, key);
		if (v !== undefined) {
			dv[key] = v;
		}
	}
	if (fromExtension) {
		dv.fromExtension = true;
	}
	return dv;
}

/** 2007 `<dataValidations>` plus the 2010 `x14:dataValidations` extension payload. */
export function parseDataValidations(root: XmlElement, extPayload?: XmlElement): DataValidation[] {
	const out: DataValidation[] = [];
	for (const dv of kids(kid(root, 'dataValidations'), 'dataValidation')) {
		const parsed = parseValidation(dv, str(dv, 'sqref') ?? '', false);
		if (parsed) {
			out.push(parsed);
		}
	}
	for (const dv of kids(extPayload, 'dataValidation')) {
		const parsed = parseValidation(dv, textContent(kid(dv, 'sqref')), true);
		if (parsed) {
			out.push(parsed);
		}
	}
	return out;
}

import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { isMultiRefValue, isRefValue, isReference } from '../../types';
import { isMissingNode } from '../logical/lift';
import { withFilePrefix } from '../logical/prefix';
import { singleArea } from './shared';

/** SHEET, SHEETS, FORMULATEXT and ISREF. */

/** SHEET([value]): 1-based index of the formula's sheet, of a reference's sheet or of a sheet named by text. */
const SHEET: FunctionDefinition = withFilePrefix({
	name: 'SHEET',
	minArgs: 0,
	maxArgs: 1,
	params: ['any'],
	impl: (args, ctx) => {
		if (args.length === 0 || isMissingNode(args[0])) {
			return ctx.sheet + 1;
		}
		const arg = args[0];
		if (isRefValue(arg)) {
			return arg.sheet + 1;
		}
		if (isMultiRefValue(arg)) {
			return arg.refs.length > 0 ? arg.refs[0].sheet + 1 : ERRORS.REF;
		}
		if (isCellError(arg)) {
			return arg;
		}
		if (typeof arg === 'string') {
			const index = ctx.sheetIndex(arg);
			return index < 0 ? ERRORS.NA : index + 1;
		}
		return ERRORS.NA;
	},
});

/** SHEETS([reference]): sheets in the workbook, or spanned by a (3-D) reference. */
const SHEETS: FunctionDefinition = withFilePrefix({
	name: 'SHEETS',
	minArgs: 0,
	maxArgs: 1,
	params: ['ref'],
	impl: (args, ctx) => {
		if (args.length === 0 || isMissingNode(args[0])) {
			return ctx.sheetCount();
		}
		const arg = args[0];
		if (isRefValue(arg)) {
			return (arg.sheetEnd ?? arg.sheet) - arg.sheet + 1;
		}
		if (isMultiRefValue(arg)) {
			return new Set(arg.refs.map((r) => r.sheet)).size;
		}
		return isCellError(arg) ? arg : ERRORS.VALUE;
	},
});

/** FORMULATEXT(reference): the formula of the top-left cell, `#N/A` when it holds a constant. */
const FORMULATEXT: FunctionDefinition = withFilePrefix({
	name: 'FORMULATEXT',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	impl: (args, ctx) => {
		const ref = singleArea(args[0]);
		if (isCellError(ref)) {
			return ref;
		}
		const formula = ctx.getFormula(ref.sheet, ref.range.top, ref.range.left);
		return formula === undefined ? ERRORS.NA : `=${formula}`;
	},
});

/** ISREF(value): the argument is a reference (never propagates errors). */
const ISREF: FunctionDefinition = {
	name: 'ISREF',
	minArgs: 1,
	maxArgs: 1,
	params: ['any'],
	propagateErrors: false,
	impl: (args) => isReference(args[0]),
};

export const SHEET_FUNCTIONS: FunctionDefinition[] = [SHEET, SHEETS, FORMULATEXT, ISREF];

import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { asArray, numberToText } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { hasValue, integerArg } from '../math/args';
import { XLFN } from '../math/define';
import { limitText } from './text-utils';

/**
 * Text form of a value for VALUETOTEXT / ARRAYTOTEXT. The strict format
 * quotes strings (doubling embedded quotes) so the result can be typed
 * back into the formula bar; numbers, booleans and errors are never quoted.
 */
export function valueAsText(v: Scalar, strict: boolean): string {
	if (typeof v === 'string') {
		return strict ? `"${v.replace(/"/g, '""')}"` : v;
	}
	if (v === null) {
		return strict ? '""' : '';
	}
	if (typeof v === 'number') {
		return numberToText(v);
	}
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	return v.code;
}

/** The `format` argument: 0 (concise, default) or 1 (strict). */
function strictFormat(args: readonly FunctionArg[], ctx: FunctionContext): boolean | CellError {
	if (!hasValue(args, 1)) {
		return false;
	}
	const f = integerArg(args[1], ctx);
	if (typeof f !== 'number') {
		return f;
	}
	return f === 0 || f === 1 ? f === 1 : ERRORS.VALUE;
}

const VALUETOTEXT: FunctionDefinition = {
	name: 'VALUETOTEXT',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	propagateErrors: false,
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const strict = strictFormat(args, ctx);
		if (isCellError(strict)) {
			return strict;
		}
		return limitText(valueAsText((args[0] ?? null) as Scalar, strict));
	},
};

const ARRAYTOTEXT: FunctionDefinition = {
	name: 'ARRAYTOTEXT',
	minArgs: 1,
	maxArgs: 2,
	params: ['array', 'value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const strict = strictFormat(args, ctx);
		if (isCellError(strict)) {
			return strict;
		}
		const array = asArray(args[0]);
		if (!strict) {
			return limitText(
				array.values
					.flat()
					.map((v) => valueAsText(v, false))
					.join(', '),
			);
		}
		const rows = array.values.map((row) => row.map((v) => valueAsText(v, true)).join(','));
		return limitText(`{${rows.join(';')}}`);
	},
};

export const TO_TEXT_FUNCTIONS: FunctionDefinition[] = [VALUETOTEXT, ARRAYTOTEXT];

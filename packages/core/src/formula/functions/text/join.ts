import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toText } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { booleanArg, isMissingArg } from '../math/args';
import { XLFN } from '../math/define';
import { limitText, textList } from './text-utils';

/**
 * Every value of an `any` argument in row-major order, blanks included
 * (references are materialised, unions visited area by area).
 */
function argValues(arg: FunctionArg, ctx: FunctionContext): Scalar[] {
	if (isMissingArg(arg)) {
		return [null];
	}
	if (isRefValue(arg)) {
		return ctx.resolveRef(arg).values.flat();
	}
	if (isMultiRefValue(arg)) {
		return arg.refs.flatMap((ref) => ctx.resolveRef(ref).values.flat());
	}
	if (isArrayValue(arg)) {
		return arg.values.flat();
	}
	if (isLambdaValue(arg)) {
		return [ERRORS.CALC];
	}
	return [arg as Scalar];
}

/** Texts of all values of the arguments, or the first error. */
function argTexts(args: readonly FunctionArg[], ctx: FunctionContext): string[] | CellError {
	const out: string[] = [];
	for (const arg of args) {
		for (const v of argValues(arg, ctx)) {
			const t = toText(v);
			if (isCellError(t)) {
				return t;
			}
			out.push(t);
		}
	}
	return out;
}

const CONCAT: FunctionDefinition = {
	name: 'CONCAT',
	minArgs: 1,
	maxArgs: 253,
	params: ['any'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const texts = argTexts(args, ctx);
		return isCellError(texts) ? texts : limitText(texts.join(''));
	},
};

/** CONCATENATE takes scalars, so a range argument lifts the function (one result per cell). */
const CONCATENATE: FunctionDefinition = {
	name: 'CONCATENATE',
	minArgs: 1,
	maxArgs: 255,
	params: ['value'],
	impl: (args) => {
		let out = '';
		for (const arg of args) {
			const t = toText(isMissingArg(arg) ? null : (arg as Scalar));
			if (isCellError(t)) {
				return t;
			}
			out += t;
		}
		return limitText(out);
	},
};

/** TEXTJOIN(delimiter, ignore_empty, text1, ...): several delimiters are used in turn. */
const TEXTJOIN: FunctionDefinition = {
	name: 'TEXTJOIN',
	minArgs: 3,
	// Delimiter and ignore_empty plus up to 252 texts.
	maxArgs: 254,
	params: ['array', 'value', 'any'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const delimiters = textList(args[0]);
		if (isCellError(delimiters)) {
			return delimiters;
		}
		const ignoreEmpty = booleanArg(args[1]);
		if (isCellError(ignoreEmpty)) {
			return ignoreEmpty;
		}
		const texts = argTexts(args.slice(2), ctx);
		if (isCellError(texts)) {
			return texts;
		}
		const items = ignoreEmpty ? texts.filter((t) => t !== '') : texts;
		const delims = delimiters.length > 0 ? delimiters : [''];
		let out = '';
		for (let i = 0; i < items.length; i++) {
			if (i > 0) {
				out += delims[(i - 1) % delims.length];
			}
			out += items[i];
			if (out.length > 32_767) {
				return ERRORS.VALUE;
			}
		}
		return out;
	},
};

export const JOIN_FUNCTIONS: FunctionDefinition[] = [CONCAT, CONCATENATE, TEXTJOIN];

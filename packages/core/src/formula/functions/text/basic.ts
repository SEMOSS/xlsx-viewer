import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition, Scalar } from '../../types';
import { isMultiRefValue, isRefValue } from '../../types';
import { integerArg, textArg } from '../math/args';
import { alias, lowerCase, upperCase } from './text-utils';

const LETTER = /\p{L}/u;
const MARK = /\p{M}/u;

/** A one-argument text function lifted over ranges. */
function textFunction(name: string, transform: (text: string) => Scalar): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		impl: (args) => {
			const text = textArg(args[0]);
			return typeof text === 'string' ? transform(text) : text;
		},
	};
}

/**
 * PROPER: upper-case the first letter of every run of letters and lower-case
 * the rest. Anything that is not a letter (digits, apostrophes, ...) starts
 * a new word, so `don't` becomes `Don'T` and `76BudGet` becomes `76Budget`.
 */
export function properCase(text: string): string {
	let out = '';
	let inWord = false;
	for (const ch of text) {
		if (LETTER.test(ch)) {
			out += inWord ? lowerCase(ch) : upperCase(ch);
			inWord = true;
		} else {
			out += ch;
			// Combining marks continue the current letter.
			inWord = inWord && MARK.test(ch);
		}
	}
	return out;
}

/** TRIM: drop leading / trailing spaces and collapse inner runs (only U+0020, not tabs or NBSP). */
export function trimSpaces(text: string): string {
	return text.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ');
}

const LEN = textFunction('LEN', (t) => t.length);

const EXACT: FunctionDefinition = {
	name: 'EXACT',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args) => {
		const a = textArg(args[0]);
		if (isCellError(a)) {
			return a;
		}
		const b = textArg(args[1]);
		return isCellError(b) ? b : a === b;
	},
};

const REPT: FunctionDefinition = {
	name: 'REPT',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		const n = integerArg(args[1], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		if (n < 0 || text.length * n > 32_767) {
			return ERRORS.VALUE;
		}
		return n === 0 || text === '' ? '' : text.repeat(n);
	},
};

/** T: the value when it is text, "" otherwise (errors propagate). */
const T: FunctionDefinition = {
	name: 'T',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args) => (typeof args[0] === 'string' ? args[0] : ''),
};

/**
 * PHONETIC: the furigana of the top-left cell of the reference. Phonetic
 * runs are not modelled, so the cell's own text is returned. Unions of
 * several areas are #N/A.
 */
const PHONETIC: FunctionDefinition = {
	name: 'PHONETIC',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	impl: (args, ctx) => {
		let arg = args[0];
		if (isMultiRefValue(arg)) {
			if (arg.refs.length !== 1) {
				return ERRORS.NA;
			}
			arg = arg.refs[0];
		}
		if (isRefValue(arg)) {
			const v = ctx.getCellValue(arg.sheet, arg.range.top, arg.range.left);
			return typeof v === 'string' || isCellError(v) ? v : '';
		}
		if (isCellError(arg)) {
			return arg;
		}
		return typeof arg === 'string' ? arg : ERRORS.NA;
	},
};

const CLEAN = textFunction('CLEAN', (t) => t.replace(/[\x00-\x1f]/g, ''));

export const BASIC_TEXT_FUNCTIONS: FunctionDefinition[] = [
	LEN,
	alias(LEN, 'LENB'),
	textFunction('LOWER', lowerCase),
	textFunction('UPPER', upperCase),
	textFunction('PROPER', properCase),
	textFunction('TRIM', trimSpaces),
	CLEAN,
	EXACT,
	REPT,
	T,
	PHONETIC,
];

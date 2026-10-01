import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { integerArg, textArg } from '../math/args';
import { XLFN } from '../math/define';
import { cp1252Char, cp1252Code, toFullWidth, toHalfWidth } from './charsets';

const CHAR: FunctionDefinition = {
	name: 'CHAR',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args, ctx) => {
		const n = integerArg(args[0], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		return n < 1 || n > 255 ? ERRORS.VALUE : cp1252Char(n);
	},
};

const CODE: FunctionDefinition = {
	name: 'CODE',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		return text === '' ? ERRORS.VALUE : cp1252Code(text.charCodeAt(0));
	},
};

const UNICHAR: FunctionDefinition = {
	name: 'UNICHAR',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const n = integerArg(args[0], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		if (n < 1 || n > 0x10ffff) {
			return ERRORS.VALUE;
		}
		// A lone surrogate half is not a character.
		if (n >= 0xd800 && n <= 0xdfff) {
			return ERRORS.NA;
		}
		return String.fromCodePoint(n);
	},
};

const UNICODE: FunctionDefinition = {
	name: 'UNICODE',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	filePrefix: XLFN,
	impl: (args) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		return text === '' ? ERRORS.VALUE : text.codePointAt(0)!;
	},
};

function widthFunction(name: string, convert: (text: string) => string): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		impl: (args) => {
			const text = textArg(args[0]);
			return isCellError(text) ? text : convert(text);
		},
	};
}

export const CHAR_FUNCTIONS: FunctionDefinition[] = [
	CHAR,
	CODE,
	UNICHAR,
	UNICODE,
	widthFunction('ASC', toHalfWidth),
	widthFunction('DBCS', toFullWidth),
];

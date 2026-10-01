import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toText } from '../../coerce';
import type { FunctionArg, FunctionDefinition, Scalar } from '../../types';
import { isArrayValue } from '../../types';
import { isMissingArg } from '../math/args';

/** Longest text a cell (and therefore a text function result) can hold. */
export const MAX_TEXT_LENGTH = 32_767;

/** The text itself, or #VALUE! when it exceeds Excel's cell limit. */
export function limitText(text: string): string | CellError {
	return text.length > MAX_TEXT_LENGTH ? ERRORS.VALUE : text;
}

function mapChars(text: string, convert: (ch: string) => string): string {
	let out = '';
	for (const ch of text) {
		const mapped = convert(ch);
		// Keep characters whose case mapping changes the length (Excel maps one to one).
		out += mapped.length === ch.length ? mapped : ch;
	}
	return out;
}

/** Upper case, character by character (`ß` stays `ß`, as in Excel). */
export function upperCase(text: string): string {
	return mapChars(text, (ch) => ch.toUpperCase());
}

/** Lower case, character by character. */
export function lowerCase(text: string): string {
	return mapChars(text, (ch) => ch.toLowerCase());
}

/** Case folding that preserves string length (and therefore match positions). */
export function foldCase(text: string): string {
	return lowerCase(text);
}

/** Escape text for literal use inside a RegExp. */
export function escapeRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
}

/**
 * Texts of an `array` parameter (delimiter lists): every element converted
 * to text, errors propagated. Blanks and explicitly empty arguments give "".
 */
export function textList(arg: FunctionArg | undefined): string[] | CellError {
	if (arg === undefined || isMissingArg(arg)) {
		return [];
	}
	const values: Scalar[] = isArrayValue(arg) ? arg.values.flat() : [arg as Scalar];
	const out: string[] = [];
	for (const v of values) {
		const t = toText(v);
		if (isCellError(t)) {
			return t;
		}
		out.push(t);
	}
	return out;
}

/** The same definition under another name (the `B` byte-count variants in a non-DBCS locale). */
export function alias(def: FunctionDefinition, name: string): FunctionDefinition {
	return { ...def, name };
}

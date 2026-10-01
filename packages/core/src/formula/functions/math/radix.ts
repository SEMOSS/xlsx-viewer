import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { integerArg, numberArg, optionalNumber, scalarArg, textArg } from './args';
import { XLFN } from './define';

const ROMAN_CHARS = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];
const MAX_SAFE = 2 ** 53;

/**
 * Roman numeral in one of Excel's five styles: 0 is classic (`CDXCIX`),
 * 1 to 4 are increasingly concise (`LDVLIV`, `XDIX`, `VDIV`, `ID` for 499).
 * The algorithm is the one used by Excel and LibreOffice.
 */
export function toRoman(value: number, form: number): string {
	let n = value;
	let out = '';
	const maxIndex = ROMAN_VALUES.length - 1;
	for (let i = 0; i <= maxIndex / 2; i++) {
		let index = 2 * i;
		const digit = Math.floor(n / ROMAN_VALUES[index]);
		if (digit % 5 === 4) {
			const index2 = digit === 4 ? index - 1 : index - 2;
			let steps = 0;
			while (steps < form && index < maxIndex) {
				steps++;
				if (ROMAN_VALUES[index2] - ROMAN_VALUES[index + 1] <= n) {
					index++;
				} else {
					steps = form;
				}
			}
			out += ROMAN_CHARS[index] + ROMAN_CHARS[index2];
			n = n + ROMAN_VALUES[index] - ROMAN_VALUES[index2];
		} else {
			if (digit > 4) {
				out += ROMAN_CHARS[index - 1];
			}
			out += ROMAN_CHARS[index].repeat(digit % 5);
			n %= ROMAN_VALUES[index];
		}
	}
	return out;
}

const ROMAN: FunctionDefinition = {
	name: 'ROMAN',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		const n = numberArg(args[0], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		const formArg = scalarArg(args, 1);
		let form: number | CellError;
		if (formArg === undefined || formArg === null || formArg === true) {
			form = 0;
		} else if (formArg === false) {
			form = 4;
		} else {
			form = integerArg(formArg, ctx);
		}
		if (typeof form !== 'number') {
			return form;
		}
		if (n < 0 || n >= 4000 || form < 0 || form > 4) {
			return ERRORS.VALUE;
		}
		return toRoman(Math.trunc(n), form);
	},
};

const ARABIC_VALUES: Record<string, number> = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };

/** Value of a Roman numeral (case-insensitive, optional leading minus), or null when invalid. */
export function fromRoman(input: string): number | null {
	let text = input.replace(/^ +| +$/g, '').toUpperCase();
	let sign = 1;
	if (text.startsWith('-')) {
		sign = -1;
		text = text.slice(1);
	}
	let total = 0;
	for (let i = 0; i < text.length; i++) {
		const v = ARABIC_VALUES[text[i]];
		if (v === undefined) {
			return null;
		}
		const next = i + 1 < text.length ? ARABIC_VALUES[text[i + 1]] : undefined;
		total += next !== undefined && v < next ? -v : v;
	}
	return total === 0 ? 0 : sign * total;
}

const ARABIC: FunctionDefinition = {
	name: 'ARABIC',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	filePrefix: XLFN,
	impl: (args) => {
		const text = textArg(args[0]);
		if (typeof text !== 'string') {
			return text;
		}
		if (text.length > 255) {
			return ERRORS.VALUE;
		}
		return fromRoman(text) ?? ERRORS.VALUE;
	},
};

const BASE: FunctionDefinition = {
	name: 'BASE',
	minArgs: 2,
	maxArgs: 3,
	params: ['value', 'value', 'value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const n = numberArg(args[0], ctx);
		if (typeof n !== 'number') {
			return n;
		}
		const radix = integerArg(args[1], ctx);
		if (typeof radix !== 'number') {
			return radix;
		}
		const minLength = optionalNumber(args, 2, ctx, 0);
		if (typeof minLength !== 'number') {
			return minLength;
		}
		const len = Math.trunc(minLength);
		if (n < 0 || n >= MAX_SAFE || radix < 2 || radix > 36 || len < 0 || len > 255) {
			return ERRORS.NUM;
		}
		return Math.trunc(n).toString(radix).toUpperCase().padStart(len, '0');
	},
};

function digitValue(code: number): number {
	if (code >= 48 && code <= 57) {
		return code - 48;
	}
	if (code >= 65 && code <= 90) {
		return code - 55;
	}
	if (code >= 97 && code <= 122) {
		return code - 87;
	}
	return 99;
}

const DECIMAL: FunctionDefinition = {
	name: 'DECIMAL',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const text = textArg(args[0]);
		if (typeof text !== 'string') {
			return text;
		}
		const radix = integerArg(args[1], ctx);
		if (typeof radix !== 'number') {
			return radix;
		}
		if (radix < 2 || radix > 36) {
			return ERRORS.NUM;
		}
		if (text.length > 255) {
			return ERRORS.VALUE;
		}
		let value = 0;
		for (let i = 0; i < text.length; i++) {
			const d = digitValue(text.charCodeAt(i));
			if (d >= radix) {
				return ERRORS.NUM;
			}
			value = value * radix + d;
		}
		return value >= MAX_SAFE ? ERRORS.NUM : value;
	},
};

export const RADIX_FUNCTIONS: FunctionDefinition[] = [ROMAN, ARABIC, BASE, DECIMAL];

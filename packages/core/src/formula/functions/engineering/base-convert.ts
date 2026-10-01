/**
 * Number base conversion: BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX,
 * DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC and OCT2HEX.
 *
 * Binary, octal and hexadecimal numbers have at most 10 digits; a
 * 10-digit number whose top bit is set is negative (two's complement over
 * 10, 30 and 40 bits). Negative results are always written with 10 digits
 * and ignore `places`.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { numberToText } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { isOmitted } from '../datetime/args';
import { engNum } from './args';

export type Radix = 2 | 8 | 10 | 16;

/** 2^(bits - 1): the first value that no longer fits (and the most negative magnitude). */
const HALF: Record<Exclude<Radix, 10>, number> = { 2: 2 ** 9, 8: 2 ** 29, 16: 2 ** 39 };
const DIGITS: Record<Exclude<Radix, 10>, RegExp> = {
	2: /^[01]*$/,
	8: /^[0-7]*$/,
	16: /^[0-9A-Fa-f]*$/,
};

/** Text of the `number` argument: numbers use Excel's general format, logicals are #VALUE!. */
function digitsArg(arg: FunctionArg | undefined): string | CellError {
	if (isOmitted(arg)) {
		return '';
	}
	if (isCellError(arg)) {
		return arg;
	}
	if (typeof arg === 'number') {
		return numberToText(arg);
	}
	return typeof arg === 'string' ? arg : ERRORS.VALUE;
}

/** Parse a binary / octal / hexadecimal number (two's complement for 10 digits). */
export function parseRadix(text: string, radix: Exclude<Radix, 10>): number | CellError {
	if (text.length > 10 || !DIGITS[radix].test(text)) {
		return ERRORS.NUM;
	}
	if (text === '') {
		return 0;
	}
	const value = parseInt(text, radix);
	return text.length === 10 && value >= HALF[radix] ? value - 2 * HALF[radix] : value;
}

/** Format a value in a radix, padded to `places` (ignored for negatives); out of range is #NUM!. */
export function formatRadix(
	value: number,
	radix: Exclude<Radix, 10>,
	places: number | undefined,
): string | CellError {
	const half = HALF[radix];
	if (value < -half || value >= half) {
		return ERRORS.NUM;
	}
	if (value < 0) {
		return (value + 2 * half).toString(radix).toUpperCase();
	}
	const text = value.toString(radix).toUpperCase();
	if (places === undefined) {
		return text;
	}
	return text.length > places ? ERRORS.NUM : text.padStart(places, '0');
}

/** The optional `places` argument: truncated, 1-10, else #NUM!. */
function placesArg(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
): number | undefined | CellError {
	if (isOmitted(arg)) {
		return undefined;
	}
	const n = engNum(arg, ctx);
	if (isCellError(n)) {
		return n;
	}
	const places = Math.trunc(n);
	return places < 1 || places > 10 ? ERRORS.NUM : places;
}

function convert(from: Radix, to: Radix) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		let value: number | CellError;
		if (from === 10) {
			const n = engNum(args[0], ctx);
			value = isCellError(n) ? n : Math.trunc(n);
		} else {
			const text = digitsArg(args[0]);
			value = isCellError(text) ? text : parseRadix(text, from);
		}
		if (isCellError(value)) {
			return value;
		}
		if (to === 10) {
			return value;
		}
		const places = placesArg(args[1], ctx);
		return isCellError(places) ? places : formatRadix(value, to, places);
	};
}

const NAMES: Record<Radix, string> = { 2: 'BIN', 8: 'OCT', 10: 'DEC', 16: 'HEX' };
const RADICES: Radix[] = [2, 8, 10, 16];

export const BASE_FUNCTIONS: FunctionDefinition[] = RADICES.flatMap((from) =>
	RADICES.filter((to) => to !== from).map(
		(to): FunctionDefinition => ({
			name: `${NAMES[from]}2${NAMES[to]}`,
			minArgs: 1,
			maxArgs: to === 10 ? 1 : 2,
			params: ['value'],
			impl: convert(from, to),
		}),
	),
);

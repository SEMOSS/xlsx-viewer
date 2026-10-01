import { formatValue } from '../../../numfmt';
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { parseNumericText } from '../../../values/parse-number';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import {
	booleanArg,
	hasValue,
	integerArg,
	isMissingArg,
	numberArg,
	scalarArg,
	textArg,
} from '../math/args';
import { XLFN } from '../math/define';
import { roundHalfAway } from '../math/rounding';

/** Largest decimal count FIXED and DOLLAR accept. */
const MAX_DECIMALS = 127;

/**
 * Plain decimal digits of a non-negative number already rounded to
 * `decimals` places, using Excel's 15 significant digits (further digits
 * are zeros), optionally with thousands separators.
 */
export function formatDecimal(abs: number, decimals: number, grouping: boolean): string {
	let intPart = '0';
	let fracPart = '';
	if (abs !== 0) {
		const [mantissa, exp] = abs.toExponential(14).split('e');
		const digits = mantissa.replace('.', '');
		const e = Number(exp);
		if (e < 0) {
			fracPart = '0'.repeat(-e - 1) + digits;
		} else if (e + 1 >= digits.length) {
			intPart = digits + '0'.repeat(e + 1 - digits.length);
		} else {
			intPart = digits.slice(0, e + 1);
			fracPart = digits.slice(e + 1);
		}
	}
	if (grouping) {
		intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
	}
	if (decimals <= 0) {
		return intPart;
	}
	return `${intPart}.${fracPart.padEnd(decimals, '0').slice(0, decimals)}`;
}

/** Shared argument handling of FIXED and DOLLAR: the rounded number and decimal count. */
function roundedArgs(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
): { value: number; decimals: number } | CellError {
	const n = numberArg(args[0], ctx);
	if (typeof n !== 'number') {
		return n;
	}
	const decimals = args.length > 1 ? integerArg(args[1], ctx) : 2;
	if (typeof decimals !== 'number') {
		return decimals;
	}
	if (decimals > MAX_DECIMALS) {
		return ERRORS.VALUE;
	}
	const value = roundHalfAway(n, Math.max(decimals, -MAX_DECIMALS));
	return { value: value === 0 ? 0 : value, decimals };
}

const FIXED: FunctionDefinition = {
	name: 'FIXED',
	minArgs: 1,
	maxArgs: 3,
	params: ['value', 'value', 'value'],
	impl: (args, ctx) => {
		const r = roundedArgs(args, ctx);
		if (isCellError(r)) {
			return r;
		}
		const noCommas = args.length > 2 ? booleanArg(args[2]) : false;
		if (typeof noCommas !== 'boolean') {
			return noCommas;
		}
		const body = formatDecimal(Math.abs(r.value), r.decimals, !noCommas);
		return r.value < 0 ? `-${body}` : body;
	},
};

const DOLLAR: FunctionDefinition = {
	name: 'DOLLAR',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		const r = roundedArgs(args, ctx);
		if (isCellError(r)) {
			return r;
		}
		const body = `$${formatDecimal(Math.abs(r.value), r.decimals, true)}`;
		return r.value < 0 ? `(${body})` : body;
	},
};

const TEXT: FunctionDefinition = {
	name: 'TEXT',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		let value = scalarArg(args, 0) ?? null;
		const format = textArg(args[1]);
		if (isCellError(format)) {
			return format;
		}
		if (format === '') {
			return '';
		}
		if (typeof value === 'boolean') {
			return value ? 'TRUE' : 'FALSE';
		}
		if (typeof value === 'string') {
			// Numeric text (numbers, dates, times, percentages) is formatted as a number.
			const parsed = parseNumericText(value, { date1904: ctx.date1904 });
			if (parsed) {
				value = parsed.value;
			}
		}
		const result = formatValue(value ?? 0, format, { date1904: ctx.date1904 });
		return result.invalid ? ERRORS.VALUE : result.text;
	},
};

const VALUE: FunctionDefinition = {
	name: 'VALUE',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args, ctx) => {
		const v = args[0];
		if (typeof v === 'number') {
			return v;
		}
		if (v === null || isMissingArg(v)) {
			return 0;
		}
		if (typeof v === 'string') {
			return parseNumericText(v, { date1904: ctx.date1904 })?.value ?? ERRORS.VALUE;
		}
		return isCellError(v) ? v : ERRORS.VALUE;
	},
};

/**
 * NUMBERVALUE's locale-independent parser: whitespace is ignored, group
 * separators may appear anywhere before the decimal separator, trailing
 * percent signs divide by 100 each. Returns null when the text is invalid.
 */
export function parseLocaleNumber(input: string, decimal: string, group: string): number | null {
	let s = input.replace(/\s+/g, '');
	if (s === '') {
		return 0;
	}
	let percents = 0;
	while (s.endsWith('%')) {
		percents++;
		s = s.slice(0, -1);
	}
	let sign = '';
	if (s.startsWith('-') || s.startsWith('+')) {
		sign = s[0] === '-' ? '-' : '';
		s = s.slice(1);
	}
	let exponent = 0;
	const exp = /[eE]([+-]?\d+)$/.exec(s);
	if (exp && decimal.toLowerCase() !== 'e' && group.toLowerCase() !== 'e') {
		exponent = Number(exp[1]);
		s = s.slice(0, exp.index);
	}
	const at = s.indexOf(decimal);
	if (at !== s.lastIndexOf(decimal)) {
		return null;
	}
	const intPart = (at < 0 ? s : s.slice(0, at)).split(group).join('');
	const fracPart = at < 0 ? '' : s.slice(at + 1);
	if (fracPart.includes(group) || !/^\d*$/.test(intPart) || !/^\d*$/.test(fracPart)) {
		return null;
	}
	if (intPart === '' && fracPart === '') {
		return null;
	}
	const value = Number(`${sign}${intPart || '0'}.${fracPart || '0'}e${exponent - 2 * percents}`);
	return Number.isFinite(value) ? value : null;
}

function separatorArg(
	args: readonly FunctionArg[],
	index: number,
	fallback: string,
): string | CellError {
	if (!hasValue(args, index)) {
		return fallback;
	}
	const t = textArg(args[index]);
	if (isCellError(t)) {
		return t;
	}
	return t === '' ? ERRORS.VALUE : t[0];
}

const NUMBERVALUE: FunctionDefinition = {
	name: 'NUMBERVALUE',
	minArgs: 1,
	maxArgs: 3,
	params: ['value', 'value', 'value'],
	filePrefix: XLFN,
	impl: (args) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		const decimal = separatorArg(args, 1, '.');
		if (isCellError(decimal)) {
			return decimal;
		}
		const group = separatorArg(args, 2, ',');
		if (isCellError(group)) {
			return group;
		}
		if (decimal === group) {
			return ERRORS.VALUE;
		}
		return parseLocaleNumber(text, decimal, group) ?? ERRORS.VALUE;
	},
};

export const FORMAT_FUNCTIONS: FunctionDefinition[] = [TEXT, FIXED, DOLLAR, VALUE, NUMBERVALUE];

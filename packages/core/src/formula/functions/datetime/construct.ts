/**
 * Building date and time serials: DATE, TIME, DATEVALUE, TIMEVALUE, NOW, TODAY.
 */
import { serialToDateParts } from '../../../values/dates';
import { ERRORS, isCellError } from '../../../values/errors';
import { parseNumericText } from '../../../values/parse-number';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { readNumbers, REQUIRED } from './args';
import { maxSerial, ymdToSerial } from './serial';

const SECONDS_PER_DAY = 86_400;
const MS_PER_DAY = 86_400_000;
/** TIME accepts each component in the 16-bit range. */
const TIME_LIMIT = 32_767;

function dateImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	let year = Math.trunc(n[0]);
	const month = Math.trunc(n[1]);
	const day = Math.trunc(n[2]);
	if (year < 0 || year >= 10_000) {
		return ERRORS.NUM;
	}
	// Two-digit style years: 0-1899 are offsets from 1900 (in both date systems).
	if (year < 1900) {
		year += 1900;
	}
	const serial = ymdToSerial(year, month, day, ctx.date1904);
	return serial < 0 || serial > maxSerial(ctx.date1904) ? ERRORS.NUM : serial;
}

function timeImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [h, m, s] = n.map(Math.trunc);
	if ([h, m, s].some((v) => v > TIME_LIMIT || v < -TIME_LIMIT - 1)) {
		return ERRORS.NUM;
	}
	const total = h * 3600 + m * 60 + s;
	if (total < 0) {
		return ERRORS.NUM;
	}
	return (total % SECONDS_PER_DAY) / SECONDS_PER_DAY;
}

/** Parse date / time text the way a cell entry would be (DATEVALUE / TIMEVALUE need real text). */
function parseDateText(arg: FunctionArg, ctx: FunctionContext) {
	if (typeof arg !== 'string') {
		return null;
	}
	const defaultYear = serialToDateParts(ctx.now(), ctx.date1904).year;
	return parseNumericText(arg, { date1904: ctx.date1904, defaultYear });
}

function dateValueImpl(args: FunctionArg[], ctx: FunctionContext) {
	const parsed = parseDateText(args[0], ctx);
	if (!parsed || (parsed.kind !== 'date' && parsed.kind !== 'datetime')) {
		return ERRORS.VALUE;
	}
	const serial = Math.floor(parsed.value);
	return serial < 0 || serial > maxSerial(ctx.date1904) ? ERRORS.VALUE : serial;
}

function timeValueImpl(args: FunctionArg[], ctx: FunctionContext) {
	const parsed = parseDateText(args[0], ctx);
	if (!parsed || (parsed.kind !== 'time' && parsed.kind !== 'datetime' && parsed.kind !== 'date')) {
		return ERRORS.VALUE;
	}
	// Keep only the time of day; rounding to the millisecond removes the noise
	// introduced by adding the time to a date serial.
	const frac = parsed.value - Math.floor(parsed.value);
	return (Math.round(frac * MS_PER_DAY) % MS_PER_DAY) / MS_PER_DAY;
}

export const CONSTRUCT_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'DATE', minArgs: 3, maxArgs: 3, params: ['value'], impl: dateImpl },
	{ name: 'TIME', minArgs: 3, maxArgs: 3, params: ['value'], impl: timeImpl },
	{ name: 'DATEVALUE', minArgs: 1, maxArgs: 1, params: ['value'], impl: dateValueImpl },
	{ name: 'TIMEVALUE', minArgs: 1, maxArgs: 1, params: ['value'], impl: timeValueImpl },
	{
		name: 'NOW',
		minArgs: 0,
		maxArgs: 0,
		params: [],
		volatile: true,
		impl: (_args, ctx) => ctx.now(),
	},
	{
		name: 'TODAY',
		minArgs: 0,
		maxArgs: 0,
		params: [],
		volatile: true,
		impl: (_args, ctx) => Math.floor(ctx.now()),
	},
];

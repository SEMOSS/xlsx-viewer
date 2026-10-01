/**
 * Reading parts of a date / time serial: YEAR, MONTH, DAY, HOUR, MINUTE,
 * SECOND, WEEKDAY, WEEKNUM and ISOWEEKNUM.
 */
import { serialToDateParts } from '../../../values/dates';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { intArg, numArg, type NumResult } from './args';
import { maxSerial, serialToYmd, weekdayOf, ymdToSerial } from './serial';

/** A serial with its time of day, validated against the date system's range. */
function serialArg(arg: FunctionArg, ctx: FunctionContext): NumResult {
	const n = numArg(arg, ctx);
	if (isCellError(n)) {
		return n;
	}
	return n < 0 || n >= maxSerial(ctx.date1904) + 1 ? ERRORS.NUM : n;
}

function datePart(part: 'year' | 'month' | 'day') {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = serialArg(args[0], ctx);
		return isCellError(n) ? n : serialToDateParts(n, ctx.date1904)[part];
	};
}

/** Time of day in whole seconds; HOUR / MINUTE / SECOND round to the nearest second. */
function timePart(part: 'hour' | 'minute' | 'second') {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = serialArg(args[0], ctx);
		if (isCellError(n)) {
			return n;
		}
		const secs = Math.round((n - Math.floor(n)) * 86_400) % 86_400;
		if (part === 'hour') {
			return Math.floor(secs / 3600);
		}
		return part === 'minute' ? Math.floor((secs % 3600) / 60) : secs % 60;
	};
}

/** First day of the week (0 = Sunday) for WEEKDAY return types 11-17 and WEEKNUM types 11-17. */
function startOfWeek(type: number): number {
	return (type - 10) % 7;
}

function weekdayImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = serialArg(args[0], ctx);
	const type = intArg(args[1], ctx, 1);
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(type)) {
		return type;
	}
	const w = weekdayOf(n, ctx.date1904);
	switch (type) {
		case 1:
			return w + 1;
		case 2:
			return ((w + 6) % 7) + 1;
		case 3:
			return (w + 6) % 7;
		default:
			if (type >= 11 && type <= 17) {
				return ((w - startOfWeek(type) + 7) % 7) + 1;
			}
			return ERRORS.NUM;
	}
}

/** ISO 8601 week number: weeks start on Monday, week 1 contains the year's first Thursday. */
function isoWeekNumber(serial: number, date1904: boolean): number {
	const day = Math.floor(serial);
	const isoDow = weekdayOf(day, date1904) || 7;
	const thursday = day - isoDow + 4;
	const year = serialToDateParts(thursday, date1904).year;
	return Math.floor((thursday - ymdToSerial(year, 1, 1, date1904)) / 7) + 1;
}

function weekNumImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = serialArg(args[0], ctx);
	const type = intArg(args[1], ctx, 1);
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(type)) {
		return type;
	}
	if (type === 21) {
		return isoWeekNumber(n, ctx.date1904);
	}
	let start: number;
	if (type === 1) {
		start = 0;
	} else if (type === 2) {
		start = 1;
	} else if (type >= 11 && type <= 17) {
		start = startOfWeek(type);
	} else {
		return ERRORS.NUM;
	}
	// Week 1 is the week containing January 1.
	const day = Math.floor(n);
	const jan1 = ymdToSerial(serialToYmd(day, ctx.date1904).year, 1, 1, ctx.date1904);
	const offset = (weekdayOf(jan1, ctx.date1904) - start + 7) % 7;
	return Math.floor((day - jan1 + offset) / 7) + 1;
}

function isoWeekNumImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = serialArg(args[0], ctx);
	return isCellError(n) ? n : isoWeekNumber(n, ctx.date1904);
}

export const COMPONENT_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'YEAR', minArgs: 1, maxArgs: 1, params: ['value'], impl: datePart('year') },
	{ name: 'MONTH', minArgs: 1, maxArgs: 1, params: ['value'], impl: datePart('month') },
	{ name: 'DAY', minArgs: 1, maxArgs: 1, params: ['value'], impl: datePart('day') },
	{ name: 'HOUR', minArgs: 1, maxArgs: 1, params: ['value'], impl: timePart('hour') },
	{ name: 'MINUTE', minArgs: 1, maxArgs: 1, params: ['value'], impl: timePart('minute') },
	{ name: 'SECOND', minArgs: 1, maxArgs: 1, params: ['value'], impl: timePart('second') },
	{ name: 'WEEKDAY', minArgs: 1, maxArgs: 2, params: ['value'], impl: weekdayImpl },
	{ name: 'WEEKNUM', minArgs: 1, maxArgs: 2, params: ['value'], impl: weekNumImpl },
	{
		name: 'ISOWEEKNUM',
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		filePrefix: '_xlfn.',
		impl: isoWeekNumImpl,
	},
];

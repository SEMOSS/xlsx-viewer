/**
 * Date arithmetic: DAYS, DAYS360, EDATE, EOMONTH, DATEDIF and YEARFRAC.
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { boolArg, intArg, readDates, textArg } from './args';
import { isBasis, yearFrac } from './day-count';
import {
	addMonths,
	isLastDayOfFeb,
	maxSerial,
	monthStartSerial,
	serialToYmd,
	ymdToSerial,
} from './serial';

function daysImpl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 2);
	return isCellError(d) ? d : d[0] - d[1];
}

/**
 * DAYS360 with Excel's actual rules. US (NASD) method: a start on the 31st
 * or on the last day of February becomes the 30th; an end on the 31st
 * becomes the 30th only when the (adjusted) start is the 30th. European
 * method: every 31st becomes the 30th.
 */
function days360Impl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 2);
	const european = boolArg(args[2], false);
	if (isCellError(d)) {
		return d;
	}
	if (isCellError(european)) {
		return european;
	}
	const a = serialToYmd(d[0], ctx.date1904);
	const b = serialToYmd(d[1], ctx.date1904);
	let sd = a.day;
	let ed = b.day;
	if (european) {
		sd = Math.min(sd, 30);
		ed = Math.min(ed, 30);
	} else {
		if (sd === 31 || isLastDayOfFeb(a, ctx.date1904)) {
			sd = 30;
		}
		if (ed === 31 && sd === 30) {
			ed = 30;
		}
	}
	return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (ed - sd);
}

/** A serial result, `#NUM!` when it falls outside the date system. */
function inRange(serial: number, ctx: FunctionContext) {
	return serial < 0 || serial > maxSerial(ctx.date1904) ? ERRORS.NUM : serial;
}

function edateImpl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 1);
	const months = intArg(args[1], ctx);
	if (isCellError(d)) {
		return d;
	}
	if (isCellError(months)) {
		return months;
	}
	return inRange(addMonths(serialToYmd(d[0], ctx.date1904), months, false, ctx.date1904), ctx);
}

function eomonthImpl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 1);
	const months = intArg(args[1], ctx);
	if (isCellError(d)) {
		return d;
	}
	if (isCellError(months)) {
		return months;
	}
	const start = serialToYmd(d[0], ctx.date1904);
	return inRange(monthStartSerial(start.year, start.month + months + 1, ctx.date1904) - 1, ctx);
}

/**
 * DATEDIF, including Excel's quirks: "MD" borrows the start's day from the
 * month before the end month (so 31-Jan to 1-Mar is -2), and "YD" moves the
 * end date into the start's year (or the year after) before subtracting.
 */
function datedifImpl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 2);
	const unitText = textArg(args[2]);
	if (isCellError(d)) {
		return d;
	}
	if (isCellError(unitText)) {
		return unitText;
	}
	const [start, end] = d;
	if (start > end) {
		return ERRORS.NUM;
	}
	const a = serialToYmd(start, ctx.date1904);
	const b = serialToYmd(end, ctx.date1904);
	const dayBefore = b.day < a.day ? 1 : 0;
	const months = (b.year - a.year) * 12 + (b.month - a.month) - dayBefore;
	switch (unitText.toUpperCase()) {
		case 'D':
			return end - start;
		case 'M':
			return months;
		case 'Y':
			return Math.floor(months / 12);
		case 'YM':
			return months % 12;
		case 'MD':
			if (b.day >= a.day) {
				return b.day - a.day;
			}
			return end - ymdToSerial(b.year, b.month - 1, a.day, ctx.date1904);
		case 'YD': {
			const endFirst = b.month > a.month || (b.month === a.month && b.day >= a.day);
			const year = endFirst ? a.year : a.year + 1;
			return ymdToSerial(year, b.month, b.day, ctx.date1904) - start;
		}
		default:
			return ERRORS.NUM;
	}
}

function yearFracImpl(args: FunctionArg[], ctx: FunctionContext) {
	const d = readDates(args, ctx, 2);
	const basis = intArg(args[2], ctx, 0);
	if (isCellError(d)) {
		return d;
	}
	if (isCellError(basis)) {
		return basis;
	}
	if (!isBasis(basis)) {
		return ERRORS.NUM;
	}
	return yearFrac(d[0], d[1], basis, ctx.date1904);
}

export const ARITHMETIC_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'DAYS', minArgs: 2, maxArgs: 2, params: ['value'], filePrefix: '_xlfn.', impl: daysImpl },
	{ name: 'DAYS360', minArgs: 2, maxArgs: 3, params: ['value'], impl: days360Impl },
	{ name: 'EDATE', minArgs: 2, maxArgs: 2, params: ['value'], impl: edateImpl },
	{ name: 'EOMONTH', minArgs: 2, maxArgs: 2, params: ['value'], impl: eomonthImpl },
	{ name: 'DATEDIF', minArgs: 3, maxArgs: 3, params: ['value'], impl: datedifImpl },
	{ name: 'YEARFRAC', minArgs: 2, maxArgs: 3, params: ['value'], impl: yearFracImpl },
];

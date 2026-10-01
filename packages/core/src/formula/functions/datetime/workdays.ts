/**
 * Working-day arithmetic: NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY and WORKDAY.INTL.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { forEachScalar, toNumber } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { intArg, isOmitted, readDates } from './args';
import { maxSerial, weekdayOf } from './serial';

/** `weekend[d]` is true when weekday `d` (0 = Sunday) is a non-working day. */
type WeekendMask = readonly boolean[];

const SAT_SUN: WeekendMask = [true, false, false, false, false, false, true];

function maskOf(days: number[]): WeekendMask {
	const mask = new Array<boolean>(7).fill(false);
	for (const d of days) {
		mask[d] = true;
	}
	return mask;
}

/**
 * The `weekend` argument of the .INTL functions: a code 1-7 (two-day
 * weekends starting Saturday/Sunday, Sunday/Monday, ...), 11-17 (single
 * days, Sunday ... Saturday) or a seven-character 0/1 string starting on
 * Monday. Omitted means Saturday + Sunday.
 */
export function parseWeekend(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
): WeekendMask | CellError {
	if (isOmitted(arg)) {
		return SAT_SUN;
	}
	if (isCellError(arg)) {
		return arg;
	}
	if (typeof arg === 'string') {
		if (!/^[01]{7}$/.test(arg)) {
			return ERRORS.VALUE;
		}
		const mask = new Array<boolean>(7).fill(false);
		for (let i = 0; i < 7; i++) {
			mask[(i + 1) % 7] = arg[i] === '1';
		}
		return mask;
	}
	if (typeof arg !== 'number') {
		return ERRORS.VALUE;
	}
	const code = Math.trunc(arg);
	if (code >= 1 && code <= 7) {
		return maskOf([(code + 5) % 7, (code + 6) % 7]);
	}
	if (code >= 11 && code <= 17) {
		return maskOf([code - 11]);
	}
	return ERRORS.NUM;
}

/** Distinct holiday serials (whole days) from a holidays argument. */
export function collectHolidays(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
): number[] | CellError {
	if (isOmitted(arg)) {
		return [];
	}
	const days = new Set<number>();
	let error: CellError | undefined;
	forEachScalar(arg as FunctionArg, ctx, (v) => {
		if (error || v === null) {
			return;
		}
		if (typeof v === 'boolean') {
			error = ERRORS.VALUE;
			return;
		}
		const n = toNumber(v, ctx.date1904);
		if (isCellError(n)) {
			error = n;
			return;
		}
		const d = Math.floor(n);
		if (d < 0 || d > maxSerial(ctx.date1904)) {
			error = ERRORS.NUM;
			return;
		}
		days.add(d);
	});
	return error ?? [...days].sort((x, y) => x - y);
}

function isWorkday(serial: number, weekend: WeekendMask, date1904: boolean): boolean {
	return !weekend[weekdayOf(serial, date1904)];
}

/** Working days between two dates inclusive; negative when `end` precedes `start`. */
export function countWorkdays(
	start: number,
	end: number,
	weekend: WeekendMask,
	holidays: readonly number[],
	date1904: boolean,
): number {
	const lo = Math.min(start, end);
	const hi = Math.max(start, end);
	const total = hi - lo + 1;
	const weeks = Math.floor(total / 7);
	const perWeek = weekend.filter((w) => !w).length;
	let count = weeks * perWeek;
	for (let d = lo + weeks * 7; d <= hi; d++) {
		if (isWorkday(d, weekend, date1904)) {
			count++;
		}
	}
	for (const h of holidays) {
		if (h >= lo && h <= hi && isWorkday(h, weekend, date1904)) {
			count--;
		}
	}
	return start <= end ? count : -count;
}

/** Step `count` working days (weekends only) from `from` in direction `dir`. */
function stepWorkdays(
	from: number,
	count: number,
	dir: 1 | -1,
	weekend: WeekendMask,
	perWeek: number,
	date1904: boolean,
): number {
	const weeks = Math.floor((count - 1) / perWeek);
	let d = from + dir * weeks * 7;
	let remaining = count - weeks * perWeek;
	while (remaining > 0) {
		d += dir;
		if (isWorkday(d, weekend, date1904)) {
			remaining--;
		}
	}
	return d;
}

/** The date `days` working days from `start`, skipping weekends and holidays (WORKDAY). */
export function addWorkdays(
	start: number,
	days: number,
	weekend: WeekendMask,
	holidays: readonly number[],
	date1904: boolean,
): number | CellError {
	if (days === 0) {
		return start;
	}
	const perWeek = weekend.filter((w) => !w).length;
	if (perWeek === 0) {
		return ERRORS.VALUE;
	}
	const dir: 1 | -1 = days > 0 ? 1 : -1;
	const limit = maxSerial(date1904);
	if (Math.abs(days) / perWeek > limit / 7 + 2) {
		return ERRORS.NUM;
	}
	const offDays = holidays.filter((h) => isWorkday(h, weekend, date1904));
	let d = stepWorkdays(start, Math.abs(days), dir, weekend, perWeek, date1904);
	// Every working-day holiday passed over pushes the result one more working day out;
	// repeat for the holidays inside each extension.
	let from = start;
	for (;;) {
		const lo = from;
		const hi = d;
		const passed = offDays.filter((h) => (dir > 0 ? h > lo && h <= hi : h < lo && h >= hi)).length;
		if (passed === 0) {
			break;
		}
		from = d;
		d = stepWorkdays(d, passed, dir, weekend, perWeek, date1904);
		if (d < -7 || d > limit + 7) {
			return ERRORS.NUM;
		}
	}
	return d < 0 || d > limit ? ERRORS.NUM : d;
}

function networkdays(intl: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const d = readDates(args, ctx, 2);
		if (isCellError(d)) {
			return d;
		}
		const weekend = intl ? parseWeekend(args[2], ctx) : SAT_SUN;
		if (isCellError(weekend)) {
			return weekend;
		}
		const holidays = collectHolidays(args[intl ? 3 : 2], ctx);
		if (isCellError(holidays)) {
			return holidays;
		}
		return countWorkdays(d[0], d[1], weekend, holidays, ctx.date1904);
	};
}

function workday(intl: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const d = readDates(args, ctx, 1);
		const days = intArg(args[1], ctx);
		if (isCellError(d)) {
			return d;
		}
		if (isCellError(days)) {
			return days;
		}
		const weekend = intl ? parseWeekend(args[2], ctx) : SAT_SUN;
		if (isCellError(weekend)) {
			return weekend;
		}
		const holidays = collectHolidays(args[intl ? 3 : 2], ctx);
		if (isCellError(holidays)) {
			return holidays;
		}
		return addWorkdays(d[0], days, weekend, holidays, ctx.date1904);
	};
}

export const WORKDAY_FUNCTIONS: FunctionDefinition[] = [
	{
		name: 'NETWORKDAYS',
		minArgs: 2,
		maxArgs: 3,
		params: ['value', 'value', 'any'],
		impl: networkdays(false),
	},
	{
		name: 'NETWORKDAYS.INTL',
		minArgs: 2,
		maxArgs: 4,
		params: ['value', 'value', 'value', 'any'],
		filePrefix: '_xlfn.',
		impl: networkdays(true),
	},
	{
		name: 'WORKDAY',
		minArgs: 2,
		maxArgs: 3,
		params: ['value', 'value', 'any'],
		impl: workday(false),
	},
	{
		name: 'WORKDAY.INTL',
		minArgs: 2,
		maxArgs: 4,
		params: ['value', 'value', 'value', 'any'],
		filePrefix: '_xlfn.',
		impl: workday(true),
	},
];

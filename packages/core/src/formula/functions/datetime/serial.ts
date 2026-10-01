/**
 * Calendar arithmetic on Excel date serials, shared by the date/time and
 * financial functions.
 *
 * Everything here works in *Excel's* calendar: in the 1900 date system
 * February 1900 has 29 days (serial 60 is the phantom 1900-02-29) and
 * serial 0 is "1900-01-00". Pure integer math is used to build serials so
 * that out-of-range intermediate years (DATE month / day overflow) never hit
 * the two-digit-year quirks of `Date.UTC`.
 */
import { isLeapYear, serialToDateParts } from '../../../values/dates';

/** Serial of 1970-01-01 in the 1900 system, before the 1900 leap-year shift. */
const EPOCH_1900 = 25_569;
/** Serial of 1970-01-01 in the 1904 system. */
const EPOCH_1904 = 24_107;

/** A calendar date as Excel sees it. */
export interface Ymd {
	year: number;
	/** 1-12 */
	month: number;
	/** 1-31 (0 for serial 0 in the 1900 system) */
	day: number;
}

/** Largest valid serial (9999-12-31) of a date system. */
export function maxSerial(date1904: boolean): number {
	return date1904 ? 2_957_003 : 2_958_465;
}

/** Days from 1970-01-01 to a proleptic Gregorian date (month 1-12, any integer year). */
function daysFromCivil(year: number, month: number, day: number): number {
	const y = month <= 2 ? year - 1 : year;
	const era = Math.floor(y / 400);
	const yoe = y - era * 400;
	const mp = month > 2 ? month - 3 : month + 9;
	const doy = Math.floor((153 * mp + 2) / 5) + day - 1;
	const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
	return era * 146_097 + doe - 719_468;
}

/** Serial of the first day of a month; `month` may be any integer and rolls over years. */
export function monthStartSerial(year: number, month: number, date1904: boolean): number {
	const total = year * 12 + (month - 1);
	const y = Math.floor(total / 12);
	const m = total - y * 12 + 1;
	const days = daysFromCivil(y, m, 1);
	if (date1904) {
		return days + EPOCH_1904;
	}
	const serial = days + EPOCH_1900;
	// Before 1900-03-01 Excel's numbering is one lower than a true day count.
	return serial < 61 ? serial - 1 : serial;
}

/**
 * Serial for a date with `DATE` overflow semantics: month and day may be out
 * of range (month 13 is next January, day 0 is the last day of the previous
 * month). Days are counted in Excel's calendar, so 1900-02-29 is serial 60.
 */
export function ymdToSerial(year: number, month: number, day: number, date1904: boolean): number {
	return monthStartSerial(year, month, date1904) + day - 1;
}

/** Calendar fields of a (whole) serial. */
export function serialToYmd(serial: number, date1904: boolean): Ymd {
	const p = serialToDateParts(Math.floor(serial), date1904);
	return { year: p.year, month: p.month, day: p.day };
}

/** Leap year in Excel's calendar (1900 counts as one in the 1900 date system). */
export function isExcelLeapYear(year: number, date1904: boolean): boolean {
	return isLeapYear(year) || (year === 1900 && !date1904);
}

const MONTH_DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/** Number of days in a month of Excel's calendar. */
export function excelDaysInMonth(year: number, month: number, date1904: boolean): number {
	if (month === 2) {
		return isExcelLeapYear(year, date1904) ? 29 : 28;
	}
	return MONTH_DAYS[month - 1];
}

/** The date is the last day of its month. */
export function isMonthEnd(date: Ymd, date1904: boolean): boolean {
	return date.day === excelDaysInMonth(date.year, date.month, date1904);
}

/** The date is the last day of February. */
export function isLastDayOfFeb(date: Ymd, date1904: boolean): boolean {
	return date.month === 2 && isMonthEnd(date, date1904);
}

/**
 * Day of the week of a serial, 0 = Sunday ... 6 = Saturday, in Excel's
 * calendar (serial 1 of the 1900 system is a Sunday, serial 0 of the 1904
 * system a Friday).
 */
export function weekdayOf(serial: number, date1904: boolean): number {
	const s = Math.floor(serial) + (date1904 ? 5 : 6);
	return ((s % 7) + 7) % 7;
}

/**
 * Move a date by whole months keeping its day of month, clamped to the
 * length of the target month; with `monthEnd` the result is always the last
 * day of the target month (EDATE / coupon schedule semantics).
 */
export function addMonths(date: Ymd, months: number, monthEnd: boolean, date1904: boolean): number {
	const total = date.year * 12 + (date.month - 1) + months;
	const y = Math.floor(total / 12);
	const m = total - y * 12 + 1;
	const dim = excelDaysInMonth(y, m, date1904);
	return ymdToSerial(y, m, monthEnd ? dim : Math.min(date.day, dim), date1904);
}

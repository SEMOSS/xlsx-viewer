/**
 * Excel date serial numbers.
 *
 * In the 1900 date system serial 1 is 1900-01-01 and Excel (for Lotus 1-2-3
 * compatibility) treats 1900 as a leap year, so serial 60 is the phantom
 * 1900-02-29 and every later serial is one day ahead of a true day count.
 * In the 1904 system serial 0 is 1904-01-01.
 *
 * All conversions here work in UTC calendar fields (no time zones): a serial
 * is a pure calendar date plus a fraction of a day.
 */

const MS_PER_DAY = 86_400_000;
/** Days from 1899-12-30 to 1970-01-01. */
const EPOCH_1900_OFFSET = 25_569;
/** Days from 1904-01-01 to 1970-01-01. */
const EPOCH_1904_OFFSET = 24_107;

export interface DateParts {
	year: number;
	/** 1-12 */
	month: number;
	/** 1-31 */
	day: number;
	hour: number;
	minute: number;
	second: number;
	millisecond: number;
	/** 0 = Sunday ... 6 = Saturday */
	weekday: number;
}

/** Serial number for a calendar date (month may overflow, e.g. month 13 = next January, like `DATE`). */
export function dateToSerial(year: number, month: number, day: number, date1904 = false): number {
	const utc = Date.UTC(year, month - 1, day);
	const days = Math.round(utc / MS_PER_DAY);
	if (date1904) {
		return days + EPOCH_1904_OFFSET;
	}
	let serial = days + EPOCH_1900_OFFSET;
	// Before 1900-03-01 the true day count is one lower than Excel's numbering.
	if (serial < 61) {
		serial -= 1;
	}
	return serial;
}

/** Serial number (with fractional time) for a JS Date's UTC fields. */
export function jsDateToSerial(date: Date, date1904 = false): number {
	const days = dateToSerial(
		date.getUTCFullYear(),
		date.getUTCMonth() + 1,
		date.getUTCDate(),
		date1904,
	);
	const ms =
		date.getUTCHours() * 3_600_000 +
		date.getUTCMinutes() * 60_000 +
		date.getUTCSeconds() * 1000 +
		date.getUTCMilliseconds();
	return days + ms / MS_PER_DAY;
}

/** Serial number for a JS Date using its *local* calendar fields (what the user sees). */
export function localDateToSerial(date: Date, date1904 = false): number {
	const days = dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), date1904);
	const ms =
		date.getHours() * 3_600_000 +
		date.getMinutes() * 60_000 +
		date.getSeconds() * 1000 +
		date.getMilliseconds();
	return days + ms / MS_PER_DAY;
}

/**
 * Break a serial number into calendar fields. Serial 0 in the 1900 system
 * is reported as 1900-01-00 and serial 60 as 1900-02-29, matching Excel.
 * Times are rounded to the nearest millisecond.
 */
export function serialToDateParts(serial: number, date1904 = false): DateParts {
	let whole = Math.floor(serial);
	let frac = serial - whole;
	let msOfDay = Math.round(frac * MS_PER_DAY);
	if (msOfDay >= MS_PER_DAY) {
		whole += 1;
		msOfDay -= MS_PER_DAY;
		frac = 0;
	}
	const hour = Math.floor(msOfDay / 3_600_000);
	const minute = Math.floor((msOfDay % 3_600_000) / 60_000);
	const second = Math.floor((msOfDay % 60_000) / 1000);
	const millisecond = msOfDay % 1000;

	if (!date1904) {
		if (whole === 0) {
			return { year: 1900, month: 1, day: 0, hour, minute, second, millisecond, weekday: 6 };
		}
		if (whole === 60) {
			return { year: 1900, month: 2, day: 29, hour, minute, second, millisecond, weekday: 3 };
		}
	}
	let days: number;
	if (date1904) {
		days = whole - EPOCH_1904_OFFSET;
	} else {
		days = whole - EPOCH_1900_OFFSET + (whole < 60 ? 1 : 0);
	}
	const d = new Date(days * MS_PER_DAY);
	return {
		year: d.getUTCFullYear(),
		month: d.getUTCMonth() + 1,
		day: d.getUTCDate(),
		hour,
		minute,
		second,
		millisecond,
		weekday: date1904 || whole > 60 ? d.getUTCDay() : (d.getUTCDay() + 6) % 7,
	};
}

/** Convert a serial number to a JS Date whose UTC fields hold the calendar value. */
export function serialToJsDate(serial: number, date1904 = false): Date {
	const p = serialToDateParts(serial, date1904);
	return new Date(Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second, p.millisecond));
}

/** Fraction of a day for a time of day. */
export function timeToSerial(hour: number, minute: number, second: number): number {
	return (hour * 3600 + minute * 60 + second) / 86_400;
}

/** Largest valid serial (9999-12-31). */
export const MAX_DATE_SERIAL = 2_958_465;

export function isLeapYear(year: number): boolean {
	return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

export function daysInMonth(year: number, month: number): number {
	return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

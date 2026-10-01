/**
 * Day-count conventions (the `basis` argument of YEARFRAC and the
 * securities functions):
 *
 * | basis | convention         |
 * | ----- | ------------------ |
 * | 0     | US (NASD) 30/360   |
 * | 1     | Actual/actual      |
 * | 2     | Actual/360         |
 * | 3     | Actual/365         |
 * | 4     | European 30/360    |
 *
 * The 30/360 US rules and the actual/actual year length follow Excel's
 * actual behaviour (as reverse-engineered for OpenFormula), not the
 * simplified descriptions in the Excel documentation.
 */
import {
	isExcelLeapYear,
	isLastDayOfFeb,
	monthStartSerial,
	serialToYmd,
	ymdToSerial,
	type Ymd,
} from './serial';

export type Basis = 0 | 1 | 2 | 3 | 4;

/** Basis argument is one of 0-4. */
export function isBasis(n: number): n is Basis {
	return n === 0 || n === 1 || n === 2 || n === 3 || n === 4;
}

function diff360(a: Ymd, sd: number, b: Ymd, ed: number): number {
	return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (ed - sd);
}

/**
 * US (NASD) 30/360 day count between two dates (start <= end).
 *
 * With `modifyBoth` false this is the rule YEARFRAC basis 0 and the
 * accrued-days counts use: an end date on the 31st only becomes the 30th
 * when the start is the 30th / 31st, and an end on the last day of February
 * only moves when the start is too. With `modifyBoth` true both ends are
 * normalised (used for the length of a coupon period).
 */
export function days360Us(a: Ymd, b: Ymd, modifyBoth: boolean, date1904: boolean): number {
	let sd = a.day;
	let ed = b.day;
	const startFebEnd = isLastDayOfFeb(a, date1904);
	if (isLastDayOfFeb(b, date1904) && (startFebEnd || modifyBoth)) {
		ed = 30;
	}
	if (ed === 31 && (sd >= 30 || modifyBoth)) {
		ed = 30;
	}
	if (sd === 31 || startFebEnd) {
		sd = 30;
	}
	return diff360(a, sd, b, ed);
}

/** European 30/360: the 31st of any month counts as the 30th. */
export function days360Eu(a: Ymd, b: Ymd): number {
	return diff360(a, Math.min(a.day, 30), b, Math.min(b.day, 30));
}

/** The dates are at most one year apart, comparing calendar fields (start <= end). */
function withinAYear(a: Ymd, b: Ymd): boolean {
	if (a.year === b.year) {
		return true;
	}
	return a.year + 1 === b.year && (a.month > b.month || (a.month === b.month && a.day >= b.day));
}

/**
 * Year length used by actual/actual (basis 1) for the span `start..end`
 * (start <= end): 366 when the span is within a year and touches a
 * February 29, 365 otherwise; for longer spans the average length of the
 * calendar years involved.
 */
export function actualYearLength(start: number, end: number, date1904: boolean): number {
	const a = serialToYmd(start, date1904);
	const b = serialToYmd(end, date1904);
	if (withinAYear(a, b)) {
		if (a.year === b.year) {
			return isExcelLeapYear(a.year, date1904) ? 366 : 365;
		}
		const feb29InStartYear =
			isExcelLeapYear(a.year, date1904) && start < ymdToSerial(a.year, 3, 1, date1904);
		const feb29InEndYear =
			isExcelLeapYear(b.year, date1904) &&
			(end >= ymdToSerial(b.year, 3, 1, date1904) || (b.month === 2 && b.day === 29));
		return feb29InStartYear || feb29InEndYear ? 366 : 365;
	}
	const years = b.year - a.year + 1;
	const days = monthStartSerial(b.year + 1, 1, date1904) - monthStartSerial(a.year, 1, date1904);
	return days / years;
}

/** Fraction of a year between two whole serials (YEARFRAC). The order of the dates does not matter. */
export function yearFrac(start: number, end: number, basis: Basis, date1904: boolean): number {
	const s = Math.min(start, end);
	const e = Math.max(start, end);
	if (s === e) {
		return 0;
	}
	switch (basis) {
		case 0:
			return days360Us(serialToYmd(s, date1904), serialToYmd(e, date1904), false, date1904) / 360;
		case 1:
			return (e - s) / actualYearLength(s, e, date1904);
		case 2:
			return (e - s) / 360;
		case 3:
			return (e - s) / 365;
		default:
			return days360Eu(serialToYmd(s, date1904), serialToYmd(e, date1904)) / 360;
	}
}

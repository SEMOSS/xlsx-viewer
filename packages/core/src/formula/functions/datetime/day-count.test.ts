import { describe, expect, it } from 'vitest';

import { actualYearLength, days360Eu, days360Us, yearFrac } from './day-count';
import {
	addMonths,
	excelDaysInMonth,
	monthStartSerial,
	serialToYmd,
	weekdayOf,
	ymdToSerial,
} from './serial';

const s = (y: number, m: number, d: number) => ymdToSerial(y, m, d, false);
const ymd = (year: number, month: number, day: number) => ({ year, month, day });

describe('serial helpers', () => {
	it('builds serials in Excel calendar', () => {
		expect(s(1900, 1, 1)).toBe(1);
		expect(s(1900, 2, 29)).toBe(60);
		expect(s(1900, 3, 1)).toBe(61);
		expect(s(2024, 1, 1)).toBe(45292);
		expect(s(2024, 13, 1)).toBe(s(2025, 1, 1));
		expect(s(2024, 0, 1)).toBe(s(2023, 12, 1));
		expect(s(2024, -11, 1)).toBe(s(2023, 1, 1));
		expect(ymdToSerial(1904, 1, 1, true)).toBe(0);
		expect(monthStartSerial(9999, 13, false)).toBe(2958466);
	});

	it('round-trips through calendar fields', () => {
		for (const serial of [1, 59, 60, 61, 366, 36526, 45292, 2958465]) {
			const p = serialToYmd(serial, false);
			expect(s(p.year, p.month, p.day)).toBe(serial);
		}
	});

	it('knows month lengths and weekdays', () => {
		expect(excelDaysInMonth(1900, 2, false)).toBe(29);
		expect(excelDaysInMonth(1900, 2, true)).toBe(28);
		expect(excelDaysInMonth(2000, 2, false)).toBe(29);
		expect(excelDaysInMonth(2100, 2, false)).toBe(28);
		expect(excelDaysInMonth(2023, 4, false)).toBe(30);
		expect(weekdayOf(45292, false)).toBe(1);
		expect(weekdayOf(45292 - 1462, true)).toBe(1);
	});

	it('adds months with clamping and month-end anchoring', () => {
		expect(addMonths(ymd(2024, 1, 31), 1, false, false)).toBe(s(2024, 2, 29));
		expect(addMonths(ymd(2023, 2, 28), 6, true, false)).toBe(s(2023, 8, 31));
		expect(addMonths(ymd(2023, 8, 30), -6, false, false)).toBe(s(2023, 2, 28));
		expect(addMonths(ymd(2023, 1, 15), -13, false, false)).toBe(s(2021, 12, 15));
	});
});

describe('30/360 day counts', () => {
	it.each([
		[[2011, 1, 31], [2011, 3, 31], false, 60],
		[[2011, 1, 30], [2011, 3, 31], false, 60],
		[[2011, 1, 29], [2011, 3, 31], false, 62],
		[[2011, 2, 28], [2011, 3, 31], false, 31],
		[[2011, 2, 28], [2012, 2, 29], false, 360],
		[[2011, 1, 31], [2011, 2, 28], false, 28],
		[[2010, 8, 31], [2011, 2, 28], false, 178],
		[[2010, 8, 31], [2011, 2, 28], true, 180],
		[[2011, 2, 28], [2011, 8, 31], true, 180],
		[[2011, 2, 28], [2011, 8, 31], false, 181],
		[[2006, 1, 1], [2006, 3, 26], false, 85],
	])('US %j -> %j (both=%s) = %d', (a, b, both, days) => {
		expect(days360Us(ymd(a[0], a[1], a[2]), ymd(b[0], b[1], b[2]), both as boolean, false)).toBe(
			days,
		);
	});

	it('European treats every 31st as the 30th', () => {
		expect(days360Eu(ymd(2011, 1, 29), ymd(2011, 3, 31))).toBe(61);
		expect(days360Eu(ymd(2011, 1, 31), ymd(2011, 2, 28))).toBe(28);
		expect(days360Eu(ymd(2011, 2, 28), ymd(2011, 3, 31))).toBe(32);
	});
});

describe('actual/actual year length', () => {
	it.each([
		[[2012, 1, 1], [2012, 7, 30], 366],
		[[2011, 1, 1], [2011, 7, 30], 365],
		[[2011, 6, 1], [2012, 3, 1], 366],
		[[2011, 6, 1], [2012, 2, 28], 365],
		[[2011, 6, 1], [2012, 2, 29], 366],
		[[2008, 2, 29], [2009, 2, 28], 366],
		[[2008, 3, 1], [2009, 3, 1], 365],
		[[2000, 1, 1], [2002, 1, 1], 1096 / 3],
		[[2010, 1, 1], [2012, 12, 31], 1096 / 3],
	])('%j -> %j uses a year of %d days', (a, b, len) => {
		expect(actualYearLength(s(a[0], a[1], a[2]), s(b[0], b[1], b[2]), false)).toBeCloseTo(len, 12);
	});
});

describe('yearFrac', () => {
	it.each([
		[[2012, 1, 1], [2012, 7, 30], 0, 0.580555555555556],
		[[2012, 1, 1], [2012, 7, 30], 1, 0.576502732240437],
		[[2012, 1, 1], [2012, 7, 30], 2, 0.586111111111111],
		[[2012, 1, 1], [2012, 7, 30], 3, 0.578082191780822],
		[[2012, 1, 1], [2012, 7, 30], 4, 0.580555555555556],
		[[2006, 1, 1], [2006, 3, 26], 1, 0.23013698630137],
		[[2000, 1, 1], [2002, 1, 1], 1, 2.000912408759124],
		[[2008, 2, 29], [2009, 2, 28], 1, 0.997267759562842],
		[[2011, 2, 28], [2012, 2, 29], 0, 1],
		[[2011, 1, 29], [2011, 3, 31], 4, 61 / 360],
	])('%j -> %j basis %d = %d', (a, b, basis, v) => {
		const start = s(a[0], a[1], a[2]);
		const end = s(b[0], b[1], b[2]);
		expect(yearFrac(start, end, basis as 0 | 1 | 2 | 3 | 4, false)).toBeCloseTo(v, 13);
		expect(yearFrac(end, start, basis as 0 | 1 | 2 | 3 | 4, false)).toBeCloseTo(v, 13);
	});
});

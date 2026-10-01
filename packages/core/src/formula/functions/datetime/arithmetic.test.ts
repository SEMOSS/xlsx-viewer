import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { DATETIME_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(DATETIME_FUNCTIONS, name), args, ctx);
const date = (y: number, m: number, d: number) => f('DATE', y, m, d) as number;

describe('DAYS', () => {
	it('subtracts whole dates', () => {
		expect(f('DAYS', '3/15/11', '2/1/11')).toBe(42);
		expect(f('DAYS', date(2011, 12, 31), date(2011, 1, 1))).toBe(364);
		expect(f('DAYS', date(2011, 1, 1), date(2011, 12, 31))).toBe(-364);
		expect(f('DAYS', 45292.9, 45292.1)).toBe(0);
		expect(f('DAYS', 61, 59)).toBe(2);
	});

	it('validates dates', () => {
		expect(f('DAYS', -1, 1)).toBe(ERRORS.NUM);
		expect(f('DAYS', 'soon', 1)).toBe(ERRORS.VALUE);
		expect(fn(DATETIME_FUNCTIONS, 'DAYS').filePrefix).toBe('_xlfn.');
	});
});

describe('DAYS360', () => {
	it.each([
		[[2011, 1, 1], [2011, 1, 30], 29],
		[[2011, 1, 1], [2011, 12, 31], 360],
		[[2011, 1, 30], [2011, 12, 31], 330],
		[[1995, 2, 28], [1995, 8, 31], 180],
		[[1996, 2, 29], [1996, 8, 31], 180],
		[[1996, 1, 30], [1996, 3, 31], 60],
		[[1996, 1, 31], [1996, 3, 31], 60],
		[[1999, 2, 2], [2000, 3, 31], 419],
		[[2000, 3, 31], [1999, 2, 2], -418],
		[[2011, 2, 28], [2011, 3, 31], 30],
		[[2011, 1, 15], [2011, 2, 28], 43],
		[[2012, 2, 28], [2012, 3, 31], 33],
	])('US method %j -> %j = %d', (a, b, days) => {
		expect(f('DAYS360', date(a[0], a[1], a[2]), date(b[0], b[1], b[2]))).toBe(days);
	});

	it.each([
		[[2011, 1, 1], [2011, 12, 31], 359],
		[[2011, 1, 31], [2011, 3, 31], 60],
		[[2011, 2, 28], [2011, 3, 31], 32],
		[[2000, 3, 31], [1999, 2, 2], -418],
	])('European method %j -> %j = %d', (a, b, days) => {
		expect(f('DAYS360', date(a[0], a[1], a[2]), date(b[0], b[1], b[2]), true)).toBe(days);
	});

	it('treats February 1900 as having 29 days', () => {
		expect(f('DAYS360', 59, 91)).toBe(33);
		expect(f('DAYS360', 60, 91)).toBe(30);
	});

	it('validates arguments', () => {
		expect(f('DAYS360', -1, 1)).toBe(ERRORS.NUM);
		expect(f('DAYS360', 1, 2, 'maybe')).toBe(ERRORS.VALUE);
		expect(f('DAYS360', 1, 31, 1)).toBe(29);
	});
});

describe('EDATE / EOMONTH', () => {
	it('shifts by whole months', () => {
		expect(f('EDATE', date(2011, 1, 15), 1)).toBe(40589);
		expect(f('EDATE', date(2011, 1, 15), -1)).toBe(40527);
		expect(f('EDATE', date(2011, 1, 15), 2)).toBe(40617);
		expect(f('EDATE', date(2011, 1, 31), 1)).toBe(date(2011, 2, 28));
		expect(f('EDATE', date(2012, 1, 31), 1)).toBe(date(2012, 2, 29));
		expect(f('EDATE', date(2024, 3, 31), -1)).toBe(date(2024, 2, 29));
		expect(f('EDATE', date(2024, 1, 31), 13)).toBe(date(2025, 2, 28));
		expect(f('EDATE', 45292.7, 1.9)).toBe(45323);
		expect(f('EDATE', '1/15/2011', '1')).toBe(40589);
	});

	it('honours the 1900 leap-year bug', () => {
		expect(f('EDATE', date(1900, 1, 31), 1)).toBe(60);
		expect(f('EOMONTH', date(1900, 1, 15), 1)).toBe(60);
		expect(f('EOMONTH', date(1900, 2, 1), 0)).toBe(60);
	});

	it('returns month ends', () => {
		expect(f('EOMONTH', date(2011, 1, 1), 1)).toBe(40602);
		expect(f('EOMONTH', date(2011, 1, 1), -3)).toBe(40482);
		expect(f('EOMONTH', date(2024, 1, 15), 0)).toBe(date(2024, 1, 31));
		expect(f('EOMONTH', date(2024, 1, 15), 1)).toBe(date(2024, 2, 29));
		expect(f('EOMONTH', date(2024, 12, 15), 2)).toBe(date(2025, 2, 28));
	});

	it('rejects results out of range', () => {
		expect(f('EDATE', date(9999, 12, 1), 1)).toBe(ERRORS.NUM);
		expect(f('EDATE', date(1900, 1, 1), -1)).toBe(ERRORS.NUM);
		expect(f('EOMONTH', date(9999, 12, 1), 1)).toBe(ERRORS.NUM);
		expect(f('EOMONTH', -1, 1)).toBe(ERRORS.NUM);
		expect(f('EDATE', 'x', 1)).toBe(ERRORS.VALUE);
	});
});

describe('DATEDIF', () => {
	const a = 37043; // 2001-06-01
	const b = 37483; // 2002-08-15

	it('matches the documented examples', () => {
		expect(f('DATEDIF', date(2001, 1, 1), date(2003, 1, 1), 'Y')).toBe(2);
		expect(f('DATEDIF', a, b, 'D')).toBe(440);
		expect(f('DATEDIF', a, b, 'YD')).toBe(75);
	});

	it('supports every unit, case-insensitively', () => {
		expect(f('DATEDIF', a, b, 'm')).toBe(14);
		expect(f('DATEDIF', a, b, 'ym')).toBe(2);
		expect(f('DATEDIF', a, b, 'md')).toBe(14);
		expect(f('DATEDIF', a, b, 'y')).toBe(1);
		expect(f('DATEDIF', date(2020, 3, 15), date(2021, 3, 10), 'Y')).toBe(0);
		expect(f('DATEDIF', date(2020, 3, 15), date(2021, 3, 10), 'M')).toBe(11);
		expect(f('DATEDIF', date(2020, 3, 15), date(2021, 3, 10), 'YM')).toBe(11);
		expect(f('DATEDIF', date(2020, 3, 15), date(2021, 3, 10), 'MD')).toBe(23);
		expect(f('DATEDIF', date(2020, 3, 15), date(2021, 3, 15), 'Y')).toBe(1);
		expect(f('DATEDIF', a, a, 'D')).toBe(0);
	});

	it('reproduces the MD and YD quirks', () => {
		expect(f('DATEDIF', date(2015, 1, 31), date(2015, 3, 1), 'MD')).toBe(-2);
		expect(f('DATEDIF', date(2011, 1, 29), date(2011, 3, 1), 'MD')).toBe(0);
		expect(f('DATEDIF', date(2012, 1, 28), date(2012, 3, 1), 'MD')).toBe(2);
		expect(f('DATEDIF', date(1988, 6, 22), date(2012, 5, 11), 'YD')).toBe(323);
		expect(f('DATEDIF', date(2007, 3, 1), date(2008, 2, 29), 'YD')).toBe(365);
	});

	it('validates arguments', () => {
		expect(f('DATEDIF', b, a, 'D')).toBe(ERRORS.NUM);
		expect(f('DATEDIF', a, b, 'W')).toBe(ERRORS.NUM);
		expect(f('DATEDIF', a, b, 1)).toBe(ERRORS.NUM);
		expect(f('DATEDIF', -1, b, 'D')).toBe(ERRORS.NUM);
	});
});

describe('YEARFRAC', () => {
	const s = 40909; // 2012-01-01
	const e = 41120; // 2012-07-30

	it('matches the documented examples', () => {
		expect(f('YEARFRAC', s, e)).toBeCloseTo(0.58055556, 8);
		expect(f('YEARFRAC', s, e, 1)).toBeCloseTo(0.57650273, 8);
		expect(f('YEARFRAC', s, e, 3)).toBeCloseTo(0.57808219, 8);
		expect(f('YEARFRAC', s, e, 2)).toBeCloseTo(211 / 360, 14);
		expect(f('YEARFRAC', s, e, 4)).toBeCloseTo(209 / 360, 14);
	});

	it('ignores the order of the dates and the time of day', () => {
		expect(f('YEARFRAC', e, s, 1)).toBe(f('YEARFRAC', s, e, 1));
		expect(f('YEARFRAC', s + 0.9, e + 0.1, 3)).toBe(211 / 365);
		expect(f('YEARFRAC', s, s, 1)).toBe(0);
	});

	it('validates the basis', () => {
		expect(f('YEARFRAC', s, e, 5)).toBe(ERRORS.NUM);
		expect(f('YEARFRAC', s, e, -1)).toBe(ERRORS.NUM);
		expect(f('YEARFRAC', s, e, 4.9)).toBeCloseTo(209 / 360, 14);
		expect(f('YEARFRAC', 'x', e)).toBe(ERRORS.VALUE);
	});
});

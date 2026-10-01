import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { DATETIME_FUNCTIONS } from './index';

const ctx = new MockContext();
const ctx1904 = Object.assign(new MockContext(), { date1904: true });
const f = (name: string, ...args: FunctionArg[]) => call(fn(DATETIME_FUNCTIONS, name), args, ctx);
const f1904 = (name: string, ...args: FunctionArg[]) =>
	call(fn(DATETIME_FUNCTIONS, name), args, ctx1904);
const date = (y: number, m: number, d: number) => f('DATE', y, m, d) as number;

describe('YEAR / MONTH / DAY', () => {
	it.each([
		[45292, 2024, 1, 1],
		[45292.75, 2024, 1, 1],
		[45351, 2024, 2, 29],
		[0, 1900, 1, 0],
		[1, 1900, 1, 1],
		[59, 1900, 2, 28],
		[60, 1900, 2, 29],
		[61, 1900, 3, 1],
		[2958465, 9999, 12, 31],
		[36526, 2000, 1, 1],
	])('serial %d is %d-%d-%d', (serial, y, m, d) => {
		expect(f('YEAR', serial)).toBe(y);
		expect(f('MONTH', serial)).toBe(m);
		expect(f('DAY', serial)).toBe(d);
	});

	it('parses date text and coerces booleans', () => {
		expect(f('YEAR', '2024-03-15')).toBe(2024);
		expect(f('MONTH', 'March 15, 2024')).toBe(3);
		expect(f('DAY', '3/15/2024')).toBe(15);
		expect(f('YEAR', true)).toBe(1900);
		expect(f('YEAR', null)).toBe(1900);
	});

	it('rejects invalid serials', () => {
		expect(f('YEAR', -1)).toBe(ERRORS.NUM);
		expect(f('MONTH', -0.5)).toBe(ERRORS.NUM);
		expect(f('DAY', 2958466)).toBe(ERRORS.NUM);
		expect(f('YEAR', 'abc')).toBe(ERRORS.VALUE);
		expect(f('YEAR', ERRORS.DIV0)).toBe(ERRORS.DIV0);
	});

	it('uses the 1904 date system', () => {
		expect(f1904('YEAR', 0)).toBe(1904);
		expect(f1904('DAY', 0)).toBe(1);
		expect(f1904('MONTH', 59)).toBe(2);
		expect(f1904('DAY', 59)).toBe(29);
		expect(f1904('YEAR', 2957004)).toBe(ERRORS.NUM);
	});

	it('lifts over arrays', () => {
		expect(f('YEAR', arr([45292], [45658]))).toEqual(arr([2024], [2025]));
	});
});

describe('HOUR / MINUTE / SECOND', () => {
	it('reads the time of day', () => {
		expect(f('HOUR', 0.75)).toBe(18);
		expect(f('MINUTE', 0.75)).toBe(0);
		expect(f('HOUR', 45292.5)).toBe(12);
		expect(f('HOUR', '3:30:30 PM')).toBe(15);
		expect(f('MINUTE', '3:30:30 PM')).toBe(30);
		expect(f('SECOND', '3:30:30 PM')).toBe(30);
		expect(f('SECOND', '4:48:18 PM')).toBe(18);
		expect(f('MINUTE', '12:45:00')).toBe(45);
		expect(f('HOUR', 0)).toBe(0);
		expect(f('SECOND', 1 - 1 / 86400)).toBe(59);
	});

	it('rounds to the nearest second', () => {
		expect(f('SECOND', 0.6 / 86400)).toBe(1);
		expect(f('SECOND', 0.4 / 86400)).toBe(0);
		const almost11 = (10 * 3600 + 59 * 60 + 59.6) / 86400;
		expect(f('HOUR', almost11)).toBe(11);
		expect(f('MINUTE', almost11)).toBe(0);
		expect(f('HOUR', 0.9999999999)).toBe(0);
	});

	it('rejects negative values and bad text', () => {
		expect(f('HOUR', -0.5)).toBe(ERRORS.NUM);
		expect(f('MINUTE', 'later')).toBe(ERRORS.VALUE);
		expect(f('SECOND', 3e6)).toBe(ERRORS.NUM);
	});
});

describe('WEEKDAY', () => {
	const thursday = 39492; // 2008-02-14

	it('supports every return type', () => {
		expect(f('WEEKDAY', thursday)).toBe(5);
		expect(f('WEEKDAY', thursday, 1)).toBe(5);
		expect(f('WEEKDAY', thursday, 2)).toBe(4);
		expect(f('WEEKDAY', thursday, 3)).toBe(3);
		expect(f('WEEKDAY', thursday, 11)).toBe(4);
		expect(f('WEEKDAY', thursday, 12)).toBe(3);
		expect(f('WEEKDAY', thursday, 13)).toBe(2);
		expect(f('WEEKDAY', thursday, 14)).toBe(1);
		expect(f('WEEKDAY', thursday, 15)).toBe(7);
		expect(f('WEEKDAY', thursday, 16)).toBe(6);
		expect(f('WEEKDAY', thursday, 17)).toBe(5);
		expect(f('WEEKDAY', thursday, 1.9)).toBe(5);
	});

	it('follows Excel calendar quirks', () => {
		expect(f('WEEKDAY', 0)).toBe(7);
		expect(f('WEEKDAY', 1)).toBe(1);
		expect(f('WEEKDAY', 60)).toBe(4);
		expect(f('WEEKDAY', 61)).toBe(5);
		expect(f('WEEKDAY', 45292)).toBe(2);
		expect(f1904('WEEKDAY', 0)).toBe(6);
	});

	it('validates arguments', () => {
		expect(f('WEEKDAY', thursday, 0)).toBe(ERRORS.NUM);
		expect(f('WEEKDAY', thursday, 4)).toBe(ERRORS.NUM);
		expect(f('WEEKDAY', thursday, 18)).toBe(ERRORS.NUM);
		expect(f('WEEKDAY', -1)).toBe(ERRORS.NUM);
		expect(f('WEEKDAY', 'x')).toBe(ERRORS.VALUE);
		expect(f('WEEKDAY', '2/14/2008')).toBe(5);
	});
});

describe('WEEKNUM / ISOWEEKNUM', () => {
	it('matches the documented examples', () => {
		expect(f('WEEKNUM', date(2012, 3, 9))).toBe(10);
		expect(f('WEEKNUM', date(2012, 3, 9), 2)).toBe(11);
		expect(f('ISOWEEKNUM', date(2012, 3, 9))).toBe(10);
	});

	it('counts weeks from the week containing January 1', () => {
		expect(f('WEEKNUM', date(2024, 1, 1))).toBe(1);
		expect(f('WEEKNUM', date(2024, 1, 6))).toBe(1);
		expect(f('WEEKNUM', date(2024, 1, 7))).toBe(2);
		expect(f('WEEKNUM', date(2024, 1, 7), 2)).toBe(1);
		expect(f('WEEKNUM', date(2024, 1, 8), 2)).toBe(2);
		expect(f('WEEKNUM', date(2024, 12, 31))).toBe(53);
		expect(f('WEEKNUM', date(2023, 1, 1), 1)).toBe(1);
		expect(f('WEEKNUM', date(2023, 1, 2), 2)).toBe(2);
		expect(f('WEEKNUM', date(2023, 1, 1), 17)).toBe(1);
		expect(f('WEEKNUM', date(2024, 1, 2), 12)).toBe(2);
		expect(f('WEEKNUM', date(2024, 1, 6), 16)).toBe(2);
		expect(f('WEEKNUM', date(2024, 1, 5), 16)).toBe(1);
	});

	it('supports the ISO system (type 21)', () => {
		expect(f('WEEKNUM', date(2021, 1, 1), 21)).toBe(53);
		expect(f('WEEKNUM', date(2024, 1, 1), 21)).toBe(1);
	});

	it.each([
		[2021, 1, 3, 53],
		[2021, 1, 4, 1],
		[2020, 12, 31, 53],
		[2019, 12, 30, 1],
		[2024, 12, 30, 1],
		[2026, 6, 15, 25],
		[2015, 12, 31, 53],
		[2016, 1, 3, 53],
	])('ISOWEEKNUM(%d-%d-%d) = %d', (y, m, d, w) => {
		expect(f('ISOWEEKNUM', date(y, m, d))).toBe(w);
	});

	it('validates arguments', () => {
		expect(f('WEEKNUM', 45292, 3)).toBe(ERRORS.NUM);
		expect(f('WEEKNUM', -1)).toBe(ERRORS.NUM);
		expect(f('ISOWEEKNUM', -1)).toBe(ERRORS.NUM);
		expect(f('ISOWEEKNUM', 'x')).toBe(ERRORS.VALUE);
	});

	it('carries the _xlfn. file prefix for ISOWEEKNUM only', () => {
		expect(fn(DATETIME_FUNCTIONS, 'ISOWEEKNUM').filePrefix).toBe('_xlfn.');
		expect(fn(DATETIME_FUNCTIONS, 'WEEKNUM').filePrefix).toBeUndefined();
	});
});

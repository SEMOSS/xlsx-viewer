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

describe('DATE', () => {
	it.each([
		[2024, 1, 1, 45292],
		[1900, 1, 1, 1],
		[1900, 2, 28, 59],
		[1900, 2, 29, 60],
		[1900, 3, 1, 61],
		[1900, 1, 0, 0],
		[1900, 3, 0, 60],
		[1900, 2, 30, 61],
		[2008, 1, 1, 39448],
		[108, 1, 2, 39449],
		[2008, 14, 2, 39846],
		[2008, -3, 2, 39327],
		[2008, 1, 35, 39482],
		[2008, 1, -15, 39432],
		[2024.9, 1.9, 1.9, 45292],
		[9999, 12, 31, 2958465],
		[2000, 0, 1, 36495],
		[1999, 12, 1, 36495],
		[2024, 2, 29, 45351],
		[2023, 2, 29, 45351 - 365],
	])('DATE(%d, %d, %d) = %d', (y, m, d, serial) => {
		expect(f('DATE', y, m, d)).toBe(serial);
	});

	it('treats years 0-1899 as offsets from 1900', () => {
		expect(f('DATE', 0, 1, 1)).toBe(1);
		expect(f('DATE', 1899, 12, 31)).toBe(f('DATE', 3799, 12, 31));
		expect(f('DATE', 24, 1, 1)).toBe(f('DATE', 1924, 1, 1));
	});

	it('rejects years and results out of range', () => {
		expect(f('DATE', 10000, 1, 1)).toBe(ERRORS.NUM);
		expect(f('DATE', -1, 1, 1)).toBe(ERRORS.NUM);
		expect(f('DATE', 9999, 12, 32)).toBe(ERRORS.NUM);
		expect(f('DATE', 1900, 1, -1)).toBe(ERRORS.NUM);
		expect(f('DATE', 2000, 1e9, 1)).toBe(ERRORS.NUM);
	});

	it('coerces text and booleans, propagates errors', () => {
		expect(f('DATE', '2024', '1', '1')).toBe(45292);
		expect(f('DATE', 2024, true, true)).toBe(45292);
		expect(f('DATE', 'x', 1, 1)).toBe(ERRORS.VALUE);
		expect(f('DATE', ERRORS.NA, 1, 1)).toBe(ERRORS.NA);
	});

	it('uses the 1904 date system', () => {
		expect(f1904('DATE', 1904, 1, 1)).toBe(0);
		expect(f1904('DATE', 2024, 1, 1)).toBe(45292 - 1462);
		expect(f1904('DATE', 1903, 12, 31)).toBe(ERRORS.NUM);
		expect(f1904('DATE', 4, 1, 2)).toBe(1);
		expect(f1904('DATE', 9999, 12, 31)).toBe(2957003);
	});

	it('lifts over arrays', () => {
		expect(f('DATE', 2024, arr([1, 2]), 1)).toEqual(arr([45292, 45323]));
	});
});

describe('TIME', () => {
	it.each([
		[12, 0, 0, 0.5],
		[18, 0, 0, 0.75],
		[0, 750, 0, 0.5208333333333334],
		[12, 30, 0, 0.5208333333333334],
		[27, 0, 0, 0.125],
		[24, 0, 0, 0],
		[1, -30, 0, 0.020833333333333332],
		[0, 0, 86400 / 4, 0.25],
		[6.9, 0.5, 0.9, 0.25],
	])('TIME(%d, %d, %d) = %d', (h, m, s, v) => {
		expect(f('TIME', h, m, s)).toBeCloseTo(v, 15);
	});

	it('matches the documented example', () => {
		expect(f('TIME', 16, 48, 10)).toBeCloseTo(0.700115740740741, 14);
	});

	it('rejects negative totals and out-of-range components', () => {
		expect(f('TIME', 0, 0, -1)).toBe(ERRORS.NUM);
		expect(f('TIME', -1, 0, 0)).toBe(ERRORS.NUM);
		expect(f('TIME', 32768, 0, 0)).toBe(ERRORS.NUM);
		expect(f('TIME', 0, 32768, 0)).toBe(ERRORS.NUM);
		expect(f('TIME', 'a', 0, 0)).toBe(ERRORS.VALUE);
	});
});

describe('DATEVALUE', () => {
	it.each([
		['8/22/2011', 40777],
		['22-MAY-2011', 40685],
		['2011/02/23', 40597],
		['1/1/2024', 45292],
		['January 1, 2024', 45292],
		['1/1/2024 10:00', 45292],
		['2/29/1900', 60],
		['5-JUL', 45478],
	])('DATEVALUE(%s) = %d', (text, serial) => {
		expect(f('DATEVALUE', text)).toBe(serial);
	});

	it('requires date text', () => {
		expect(f('DATEVALUE', 45292)).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', true)).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', null)).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', 'abc')).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', '45000')).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', '12:00')).toBe(ERRORS.VALUE);
		expect(f('DATEVALUE', '2/30/2024')).toBe(ERRORS.VALUE);
	});

	it('uses the 1904 date system', () => {
		expect(f1904('DATEVALUE', '1/2/1904')).toBe(1);
		expect(f1904('DATEVALUE', '1/1/1903')).toBe(ERRORS.VALUE);
	});
});

describe('TIMEVALUE', () => {
	it.each([
		['2:24 AM', 0.1],
		['12:00 PM', 0.5],
		['18:00', 0.75],
		['6:00:30', 0.25 + 30 / 86400],
		['1/1/2024', 0],
		['25:00', 1 / 24],
	])('TIMEVALUE(%s) = %d', (text, v) => {
		expect(f('TIMEVALUE', text)).toBeCloseTo(v, 15);
	});

	it('drops the date part of date-time text', () => {
		expect(f('TIMEVALUE', '22-Aug-2011 6:35 AM')).toBe(23700 / 86400);
	});

	it('requires time text', () => {
		expect(f('TIMEVALUE', 0.5)).toBe(ERRORS.VALUE);
		expect(f('TIMEVALUE', 'noon')).toBe(ERRORS.VALUE);
	});
});

describe('NOW / TODAY', () => {
	it('reads the context clock', () => {
		expect(f('NOW')).toBe(45292.5);
		expect(f('TODAY')).toBe(45292);
	});

	it('are volatile and take no arguments', () => {
		expect(fn(DATETIME_FUNCTIONS, 'NOW').volatile).toBe(true);
		expect(fn(DATETIME_FUNCTIONS, 'TODAY').volatile).toBe(true);
		expect(f('NOW', 1)).toBe(ERRORS.VALUE);
	});
});

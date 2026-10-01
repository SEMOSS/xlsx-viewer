import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { DATETIME_FUNCTIONS } from '../datetime';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);
const num = (name: string, ...args: FunctionArg[]) => f(name, ...args) as number;
const date = (y: number, m: number, d: number) =>
	call(fn(DATETIME_FUNCTIONS, 'DATE'), [y, m, d], ctx) as number;

describe('PRICE / YIELD', () => {
	it('match the documented examples', () => {
		expect(f('PRICE', date(2008, 2, 15), date(2017, 11, 15), 0.0575, 0.065, 100, 2, 0)).toBeCloseTo(
			94.6343616213221,
			9,
		);
		expect(
			f('YIELD', date(2008, 2, 15), date(2016, 11, 15), 0.0575, 95.04287, 100, 2, 0),
		).toBeCloseTo(0.065, 7);
	});

	it.each([
		[0, 2],
		[1, 2],
		[2, 4],
		[3, 1],
		[4, 2],
		[1, 1],
	])('round-trips with basis %d and frequency %d', (basis, freq) => {
		const s = date(2010, 3, 17);
		const m = date(2024, 9, 30);
		const price = num('PRICE', s, m, 0.045, 0.0612, 105, freq, basis);
		expect(f('YIELD', s, m, 0.045, price, 105, freq, basis)).toBeCloseTo(0.0612, 11);
	});

	it('uses simple interest in the final coupon period', () => {
		const s = date(2008, 2, 15);
		const m = date(2008, 6, 30);
		const expected = 102.875 / (1 + (135 / 180) * 0.0325) - (2.875 * 45) / 180;
		expect(f('PRICE', s, m, 0.0575, 0.065, 100, 2, 0)).toBeCloseTo(expected, 12);
		expect(f('YIELD', s, m, 0.0575, expected, 100, 2, 0)).toBeCloseTo(0.065, 12);
	});

	it('prices a zero-coupon bond and a bond at par', () => {
		const s = date(2020, 1, 1);
		const m = date(2030, 1, 1);
		expect(f('PRICE', s, m, 0, 0.05, 100, 1, 1)).toBeCloseTo(100 / Math.pow(1.05, 10), 10);
		expect(f('PRICE', s, m, 0.05, 0.05, 100, 2, 0)).toBeCloseTo(100, 10);
		expect(f('YIELD', s, m, 0.05, 100, 100, 2, 0)).toBeCloseTo(0.05, 11);
	});

	it('validate arguments', () => {
		const s = date(2008, 2, 15);
		const m = date(2017, 11, 15);
		expect(f('PRICE', s, m, -0.01, 0.065, 100, 2)).toBe(ERRORS.NUM);
		expect(f('PRICE', s, m, 0.0575, -0.065, 100, 2)).toBe(ERRORS.NUM);
		expect(f('PRICE', s, m, 0.0575, 0.065, 0, 2)).toBe(ERRORS.NUM);
		expect(f('PRICE', s, m, 0.0575, 0.065, 100, 3)).toBe(ERRORS.NUM);
		expect(f('PRICE', m, s, 0.0575, 0.065, 100, 2)).toBe(ERRORS.NUM);
		expect(f('PRICE', 'x', m, 0.0575, 0.065, 100, 2)).toBe(ERRORS.VALUE);
		expect(f('YIELD', s, m, 0.0575, 0, 100, 2)).toBe(ERRORS.NUM);
		expect(f('YIELD', s, m, 0.0575, 95, -1, 2)).toBe(ERRORS.NUM);
		expect(f('YIELD', s, m, 0.0575, 95, 100, 2, 7)).toBe(ERRORS.NUM);
	});
});

describe('DURATION / MDURATION', () => {
	it('match the documented examples', () => {
		expect(f('DURATION', date(2018, 7, 1), date(2048, 1, 1), 0.08, 0.09, 2, 1)).toBeCloseTo(
			10.9191452815919,
			9,
		);
		expect(f('MDURATION', date(2008, 1, 1), date(2016, 1, 1), 0.08, 0.09, 2, 1)).toBeCloseTo(
			5.73566981391884,
			9,
		);
	});

	it('equal the maturity of a zero-coupon bond', () => {
		expect(f('DURATION', date(2020, 1, 1), date(2025, 1, 1), 0, 0.05, 1, 1)).toBeCloseTo(5, 12);
		expect(f('MDURATION', date(2020, 1, 1), date(2025, 1, 1), 0, 0.05, 1, 1)).toBeCloseTo(
			5 / 1.05,
			12,
		);
	});

	it('relate modified to Macaulay duration', () => {
		const args = [date(2010, 5, 3), date(2031, 2, 28), 0.06, 0.071, 4, 0];
		expect(num('MDURATION', ...args)).toBeCloseTo(num('DURATION', ...args) / (1 + 0.071 / 4), 12);
	});

	it('validate arguments', () => {
		const s = date(2018, 7, 1);
		const m = date(2048, 1, 1);
		expect(f('DURATION', s, m, -0.08, 0.09, 2)).toBe(ERRORS.NUM);
		expect(f('DURATION', s, m, 0.08, -0.09, 2)).toBe(ERRORS.NUM);
		expect(f('MDURATION', s, m, 0.08, 0.09, 5)).toBe(ERRORS.NUM);
		expect(f('MDURATION', m, s, 0.08, 0.09, 2)).toBe(ERRORS.NUM);
	});
});

describe('ACCRINT / ACCRINTM', () => {
	const firstInterest = date(2008, 8, 31);
	const settlement = date(2008, 5, 1);

	it('match the documented examples', () => {
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, settlement, 0.1, 1000, 2, 0)).toBeCloseTo(
			16.6666666666667,
			10,
		);
		expect(
			f('ACCRINT', date(2008, 3, 5), firstInterest, settlement, 0.1, 1000, 2, 0, false),
		).toBeCloseTo(15.5555555555556, 10);
		expect(
			f('ACCRINT', date(2008, 4, 5), firstInterest, settlement, 0.1, 1000, 2, 0, true),
		).toBeCloseTo(7.22222222222222, 10);
		expect(f('ACCRINTM', date(2008, 4, 1), date(2008, 6, 15), 0.1, 1000, 3)).toBeCloseTo(
			20.5479452054795,
			10,
		);
	});

	it('sums quasi-coupon periods with actual day counts', () => {
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, settlement, 0.1, 1000, 2, 1)).toBeCloseTo(
			(50 * 61) / 184,
			12,
		);
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, settlement, 0.1, 1000, 2, 3)).toBeCloseTo(
			(100 * 61) / 365,
			12,
		);
	});

	it('accrues across several coupon periods', () => {
		const later = date(2009, 5, 1);
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, later, 0.1, 1000, 2, 0)).toBeCloseTo(
			50 * (2 + 61 / 180),
			10,
		);
		expect(
			f('ACCRINT', date(2008, 3, 1), firstInterest, later, 0.1, 1000, 2, 0, false),
		).toBeCloseTo(50 * (1 + 61 / 180), 10);
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, later, 0.1, 1000, 4, 0)).toBeCloseTo(
			f('ACCRINT', date(2008, 3, 1), firstInterest, later, 0.1, 1000, 2, 0) as number,
			10,
		);
	});

	it('defaults par to 1000', () => {
		expect(f('ACCRINT', date(2008, 3, 1), firstInterest, settlement, 0.1, null, 2, 0)).toBeCloseTo(
			16.6666666666667,
			10,
		);
		expect(f('ACCRINTM', date(2008, 4, 1), date(2008, 6, 15), 0.1)).toBeCloseTo(
			100 * (74 / 360),
			12,
		);
	});

	it('validate arguments', () => {
		const issue = date(2008, 3, 1);
		expect(f('ACCRINT', issue, firstInterest, settlement, 0, 1000, 2)).toBe(ERRORS.NUM);
		expect(f('ACCRINT', issue, firstInterest, settlement, 0.1, -5, 2)).toBe(ERRORS.NUM);
		expect(f('ACCRINT', issue, firstInterest, settlement, 0.1, 1000, 3)).toBe(ERRORS.NUM);
		expect(f('ACCRINT', issue, firstInterest, settlement, 0.1, 1000, 2, 5)).toBe(ERRORS.NUM);
		expect(f('ACCRINT', settlement, firstInterest, issue, 0.1, 1000, 2)).toBe(ERRORS.NUM);
		expect(f('ACCRINT', 'x', firstInterest, settlement, 0.1, 1000, 2)).toBe(ERRORS.VALUE);
		expect(f('ACCRINTM', date(2008, 6, 15), date(2008, 4, 1), 0.1)).toBe(ERRORS.NUM);
		expect(f('ACCRINTM', date(2008, 4, 1), date(2008, 6, 15), 0)).toBe(ERRORS.NUM);
		expect(f('ACCRINTM', date(2008, 4, 1), date(2008, 6, 15), 0.1, 1000, 9)).toBe(ERRORS.NUM);
	});
});

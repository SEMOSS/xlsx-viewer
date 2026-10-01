import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { DATETIME_FUNCTIONS } from '../datetime';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);
const date = (y: number, m: number, d: number) =>
	call(fn(DATETIME_FUNCTIONS, 'DATE'), [y, m, d], ctx) as number;

describe('COUP* documented examples', () => {
	const settlement = date(2011, 1, 25);
	const maturity = date(2011, 11, 15);

	it('match Excel', () => {
		expect(f('COUPDAYBS', settlement, maturity, 2, 1)).toBe(71);
		expect(f('COUPDAYS', settlement, maturity, 2, 1)).toBe(181);
		expect(f('COUPDAYSNC', settlement, maturity, 2, 1)).toBe(110);
		expect(f('COUPNCD', settlement, maturity, 2, 1)).toBe(40678);
		expect(f('COUPPCD', settlement, maturity, 2, 1)).toBe(40497);
		expect(f('COUPNUM', date(2007, 1, 25), date(2008, 11, 15), 2, 1)).toBe(4);
	});

	it.each([
		[0, 70, 180, 110],
		[1, 71, 181, 110],
		[2, 71, 180, 110],
		[3, 71, 182.5, 110],
		[4, 70, 180, 110],
	])('basis %d: days from coupon %d, period %d, to coupon %d', (basis, bs, days, nc) => {
		expect(f('COUPDAYBS', settlement, maturity, 2, basis)).toBe(bs);
		expect(f('COUPDAYS', settlement, maturity, 2, basis)).toBe(days);
		expect(f('COUPDAYSNC', settlement, maturity, 2, basis)).toBe(nc);
	});

	it('default to basis 0', () => {
		expect(f('COUPDAYBS', settlement, maturity, 2)).toBe(70);
		expect(f('COUPDAYS', settlement, maturity, 2)).toBe(180);
	});
});

describe('coupon schedules', () => {
	it('keep month ends when maturity is a month end', () => {
		const maturity = date(2011, 8, 31);
		expect(f('COUPPCD', date(2011, 3, 15), maturity, 2)).toBe(date(2011, 2, 28));
		expect(f('COUPNCD', date(2011, 3, 15), maturity, 2)).toBe(maturity);
		expect(f('COUPPCD', date(2011, 3, 15), maturity, 4)).toBe(date(2011, 2, 28));
		expect(f('COUPNCD', date(2011, 3, 15), maturity, 4)).toBe(date(2011, 5, 31));
		expect(f('COUPPCD', date(2012, 1, 15), date(2012, 8, 31), 2)).toBe(date(2011, 8, 31));
		expect(f('COUPNCD', date(2012, 1, 15), date(2012, 8, 31), 2)).toBe(date(2012, 2, 29));
	});

	it('keep the maturity day otherwise, clamped to short months', () => {
		const maturity = date(2011, 8, 30);
		expect(f('COUPPCD', date(2011, 3, 15), maturity, 2)).toBe(date(2011, 2, 28));
		expect(f('COUPNCD', date(2011, 3, 15), maturity, 2)).toBe(maturity);
		expect(f('COUPPCD', date(2011, 2, 1), maturity, 2)).toBe(date(2010, 8, 30));
		expect(f('COUPNCD', date(2011, 2, 1), maturity, 2)).toBe(date(2011, 2, 28));
	});

	it('treat a settlement on a coupon date as the start of a period', () => {
		const maturity = date(2011, 11, 15);
		expect(f('COUPDAYBS', date(2011, 5, 15), maturity, 2, 1)).toBe(0);
		expect(f('COUPPCD', date(2011, 5, 15), maturity, 2, 1)).toBe(date(2011, 5, 15));
		expect(f('COUPNCD', date(2011, 5, 15), maturity, 2, 1)).toBe(maturity);
		expect(f('COUPNUM', date(2011, 5, 15), maturity, 2, 1)).toBe(1);
		expect(f('COUPDAYSNC', date(2011, 5, 15), maturity, 2, 1)).toBe(184);
	});

	it('count coupons for every frequency', () => {
		expect(f('COUPNUM', date(2011, 1, 25), date(2015, 11, 15), 1)).toBe(5);
		expect(f('COUPNUM', date(2011, 1, 25), date(2011, 11, 15), 4)).toBe(4);
		expect(f('COUPNUM', date(2008, 2, 15), date(2017, 11, 15), 2)).toBe(20);
		expect(f('COUPNUM', date(2011, 11, 14), date(2011, 11, 15), 2)).toBe(1);
	});

	it('count actual days across February 29', () => {
		expect(f('COUPDAYS', date(2012, 1, 25), date(2012, 11, 15), 2, 1)).toBe(182);
		expect(f('COUPDAYS', date(2012, 6, 25), date(2012, 11, 15), 2, 1)).toBe(184);
	});
});

describe('COUP* validation', () => {
	const s = date(2011, 1, 25);
	const m = date(2011, 11, 15);

	it('rejects bad frequencies, bases and date orders', () => {
		expect(f('COUPNUM', s, m, 3)).toBe(ERRORS.NUM);
		expect(f('COUPNUM', s, m, 0)).toBe(ERRORS.NUM);
		expect(f('COUPNUM', s, m, 2.9)).toBe(f('COUPNUM', s, m, 2));
		expect(f('COUPDAYS', s, m, 2, 5)).toBe(ERRORS.NUM);
		expect(f('COUPDAYS', s, m, 2, -1)).toBe(ERRORS.NUM);
		expect(f('COUPPCD', m, m, 2)).toBe(ERRORS.NUM);
		expect(f('COUPPCD', m, s, 2)).toBe(ERRORS.NUM);
	});

	it('reports invalid dates as #VALUE!', () => {
		expect(f('COUPNCD', 'not a date', m, 2)).toBe(ERRORS.VALUE);
		expect(f('COUPNCD', -1, m, 2)).toBe(ERRORS.VALUE);
		expect(f('COUPNCD', s, 3e6, 2)).toBe(ERRORS.VALUE);
		expect(f('COUPNCD', '1/25/2011', '11/15/2011', 2)).toBe(40678);
		expect(f('COUPNCD', ERRORS.NA, m, 2)).toBe(ERRORS.NA);
	});
});

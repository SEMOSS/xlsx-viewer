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

describe('ODDFPRICE / ODDFYIELD', () => {
	const settlement = date(2008, 11, 11);
	const maturity = date(2021, 3, 1);
	const issue = date(2008, 10, 15);
	const first = date(2009, 3, 1);

	it('match the documented examples', () => {
		expect(
			f('ODDFPRICE', settlement, maturity, issue, first, 0.0785, 0.0625, 100, 2, 1),
		).toBeCloseTo(113.597717474079, 9);
		expect(f('ODDFYIELD', settlement, maturity, issue, first, 0.0575, 84.5, 100, 2, 0)).toBeCloseTo(
			0.0772455415978,
			10,
		);
	});

	it('equal PRICE when the first period is a regular one', () => {
		const regularIssue = date(2008, 9, 1);
		// Bases 2 and 3 count the odd period in actual days over a 360 / 365-day period, so they differ.
		for (const basis of [0, 1, 4]) {
			const odd = num(
				'ODDFPRICE',
				settlement,
				maturity,
				regularIssue,
				first,
				0.0785,
				0.0625,
				100,
				2,
				basis,
			);
			expect(odd).toBeCloseTo(num('PRICE', settlement, maturity, 0.0785, 0.0625, 100, 2, basis), 9);
		}
	});

	it('price long first periods continuously', () => {
		const quasiStart = date(2008, 9, 1);
		const atStart = num(
			'ODDFPRICE',
			settlement,
			maturity,
			quasiStart,
			first,
			0.0785,
			0.0625,
			100,
			2,
			1,
		);
		const dayBefore = num(
			'ODDFPRICE',
			settlement,
			maturity,
			quasiStart - 1,
			first,
			0.0785,
			0.0625,
			100,
			2,
			1,
		);
		expect(Math.abs(atStart - dayBefore)).toBeLessThan(0.01);
		const long = num(
			'ODDFPRICE',
			date(2009, 2, 1),
			maturity,
			date(2008, 1, 15),
			first,
			0.0785,
			0.0625,
			100,
			2,
			0,
		);
		expect(long).toBeGreaterThan(110);
		expect(long).toBeLessThan(120);
	});

	it('round-trips yields for short and long first periods', () => {
		for (const [iss, set] of [
			[issue, settlement],
			[date(2008, 1, 15), date(2008, 6, 20)],
			[date(2008, 1, 15), date(2009, 2, 1)],
		]) {
			for (const basis of [0, 1, 3]) {
				const price = num('ODDFPRICE', set, maturity, iss, first, 0.06, 0.071, 100, 2, basis);
				expect(f('ODDFYIELD', set, maturity, iss, first, 0.06, price, 100, 2, basis)).toBeCloseTo(
					0.071,
					10,
				);
			}
		}
	});

	it('validate arguments', () => {
		expect(f('ODDFPRICE', settlement, maturity, settlement, first, 0.07, 0.06, 100, 2)).toBe(
			ERRORS.NUM,
		);
		expect(f('ODDFPRICE', settlement, maturity, issue, settlement, 0.07, 0.06, 100, 2)).toBe(
			ERRORS.NUM,
		);
		expect(f('ODDFPRICE', settlement, first, issue, first, 0.07, 0.06, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDFPRICE', settlement, maturity, issue, first, -0.07, 0.06, 100, 2)).toBe(
			ERRORS.NUM,
		);
		expect(f('ODDFPRICE', settlement, maturity, issue, first, 0.07, -0.06, 100, 2)).toBe(
			ERRORS.NUM,
		);
		expect(f('ODDFPRICE', settlement, maturity, issue, first, 0.07, 0.06, 0, 2)).toBe(ERRORS.NUM);
		expect(f('ODDFPRICE', settlement, maturity, issue, first, 0.07, 0.06, 100, 3)).toBe(ERRORS.NUM);
		expect(f('ODDFPRICE', settlement, maturity, issue, first, 0.07, 0.06, 100, 2, 5)).toBe(
			ERRORS.NUM,
		);
		expect(f('ODDFYIELD', settlement, maturity, issue, first, 0.07, 0, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDFYIELD', 'x', maturity, issue, first, 0.07, 90, 100, 2)).toBe(ERRORS.VALUE);
	});
});

describe('ODDLPRICE / ODDLYIELD', () => {
	it('match the documented examples', () => {
		expect(
			f(
				'ODDLPRICE',
				date(2008, 2, 7),
				date(2008, 6, 15),
				date(2007, 10, 15),
				0.0375,
				0.0405,
				100,
				2,
				0,
			),
		).toBeCloseTo(99.8782860147213, 10);
		expect(
			f(
				'ODDLYIELD',
				date(2008, 4, 20),
				date(2008, 6, 15),
				date(2007, 12, 24),
				0.0375,
				99.875,
				100,
				2,
				0,
			),
		).toBeCloseTo(0.0451922356291692, 12);
	});

	it('round-trip between price and yield', () => {
		const args = [date(2008, 2, 7), date(2008, 6, 15), date(2007, 10, 15)];
		for (const basis of [0, 1, 2, 3, 4]) {
			const price = num('ODDLPRICE', ...args, 0.0375, 0.0405, 100, 2, basis);
			expect(f('ODDLYIELD', ...args, 0.0375, price, 100, 2, basis)).toBeCloseTo(0.0405, 12);
		}
	});

	it('handle a last period shorter than one coupon', () => {
		const price = num(
			'ODDLPRICE',
			date(2008, 4, 20),
			date(2008, 6, 15),
			date(2007, 12, 24),
			0.0375,
			0.045,
			100,
			2,
			1,
		);
		const accrued = (1.875 * (date(2008, 4, 20) - date(2007, 12, 24))) / 183;
		const total = 100 + (1.875 * (date(2008, 6, 15) - date(2007, 12, 24))) / 183;
		const dsc = (date(2008, 6, 15) - date(2008, 4, 20)) / 183;
		expect(price).toBeCloseTo(total / (1 + (dsc * 0.045) / 2) - accrued, 12);
	});

	it('validate arguments', () => {
		const s = date(2008, 2, 7);
		const m = date(2008, 6, 15);
		const last = date(2007, 10, 15);
		expect(f('ODDLPRICE', last, m, s, 0.0375, 0.0405, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDLPRICE', m, s, last, 0.0375, 0.0405, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDLPRICE', s, m, last, -0.0375, 0.0405, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDLPRICE', s, m, last, 0.0375, -0.0405, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDLYIELD', s, m, last, 0.0375, 0, 100, 2)).toBe(ERRORS.NUM);
		expect(f('ODDLYIELD', s, m, last, 0.0375, 99, 100, 8)).toBe(ERRORS.NUM);
		expect(f('ODDLYIELD', s, m, last, 0.0375, 99, 100, 2, 6)).toBe(ERRORS.NUM);
	});
});

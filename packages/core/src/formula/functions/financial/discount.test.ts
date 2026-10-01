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
const yearfrac = (s: number, e: number, b: number) =>
	call(fn(DATETIME_FUNCTIONS, 'YEARFRAC'), [s, e, b], ctx) as number;

describe('discount securities', () => {
	it('match the documented examples', () => {
		expect(f('INTRATE', date(2008, 2, 15), date(2008, 5, 15), 1000000, 1014420, 2)).toBeCloseTo(
			0.05768,
			14,
		);
		expect(f('RECEIVED', date(2008, 2, 15), date(2008, 5, 15), 1000000, 0.0575, 2)).toBeCloseTo(
			1014584.6544071,
			6,
		);
		expect(f('PRICEDISC', date(2008, 2, 16), date(2008, 3, 1), 0.0525, 100, 2)).toBeCloseTo(
			99.7958333333333,
			10,
		);
		expect(f('YIELDDISC', date(2008, 2, 16), date(2008, 3, 1), 99.795, 100, 2)).toBeCloseTo(
			0.0528225719868583,
			13,
		);
		expect(
			f('PRICEMAT', date(2008, 2, 15), date(2008, 4, 13), date(2007, 11, 11), 0.061, 0.061, 0),
		).toBeCloseTo(99.984498875557, 9);
		expect(
			f('YIELDMAT', date(2008, 3, 15), date(2008, 11, 3), date(2007, 11, 8), 0.0625, 100.0123, 0),
		).toBeCloseTo(0.0609543336915387, 12);
	});

	it('DISC divides the discount by the year fraction', () => {
		expect(f('DISC', date(2008, 2, 16), date(2008, 3, 1), 99.795, 100, 2)).toBeCloseTo(
			(0.00205 * 360) / 14,
			14,
		);
		const s = date(2018, 7, 1);
		const m = date(2048, 1, 1);
		expect(f('DISC', s, m, 97.975, 100, 1)).toBeCloseTo(0.02025 / yearfrac(s, m, 1), 15);
		expect(f('DISC', s, m, 97.975, 100)).toBeCloseTo(0.02025 / yearfrac(s, m, 0), 15);
	});

	it('round-trips prices and rates', () => {
		const s = date(2023, 1, 10);
		const m = date(2023, 10, 31);
		for (const basis of [0, 1, 2, 3, 4]) {
			const price = num('PRICEDISC', s, m, 0.043, 100, basis);
			expect(f('DISC', s, m, price, 100, basis)).toBeCloseTo(0.043, 14);
			const received = num('RECEIVED', s, m, 1000, 0.043, basis);
			expect(f('INTRATE', s, m, 1000, received, basis)).toBeCloseTo(
				num('YIELDDISC', s, m, price, 100, basis),
				12,
			);
			const issue = date(2022, 6, 1);
			const pm = num('PRICEMAT', s, m, issue, 0.05, 0.061, basis);
			expect(f('YIELDMAT', s, m, issue, 0.05, pm, basis)).toBeCloseTo(0.061, 12);
		}
	});

	it('validate arguments', () => {
		const s = date(2008, 2, 16);
		const m = date(2008, 3, 1);
		expect(f('DISC', s, m, 0, 100)).toBe(ERRORS.NUM);
		expect(f('DISC', s, m, 99, 0)).toBe(ERRORS.NUM);
		expect(f('DISC', m, s, 99, 100)).toBe(ERRORS.NUM);
		expect(f('DISC', s, s, 99, 100)).toBe(ERRORS.NUM);
		expect(f('DISC', s, m, 99, 100, 5)).toBe(ERRORS.NUM);
		expect(f('DISC', 'x', m, 99, 100)).toBe(ERRORS.VALUE);
		expect(f('INTRATE', s, m, -1, 100)).toBe(ERRORS.NUM);
		expect(f('RECEIVED', s, m, 1000, 0)).toBe(ERRORS.NUM);
		expect(f('RECEIVED', s, m, 1000, 100)).toBe(ERRORS.NUM);
		expect(f('PRICEDISC', s, m, 0, 100)).toBe(ERRORS.NUM);
		expect(f('YIELDDISC', s, m, 99, -100)).toBe(ERRORS.NUM);
		expect(f('PRICEMAT', s, m, date(2007, 1, 1), -0.01, 0.05)).toBe(ERRORS.NUM);
		expect(f('PRICEMAT', s, m, date(2009, 1, 1), 0.05, 0.05)).toBe(ERRORS.NUM);
		expect(f('YIELDMAT', s, m, date(2007, 1, 1), 0.05, 0)).toBe(ERRORS.NUM);
		expect(f('YIELDMAT', s, m, 'x', 0.05, 99)).toBe(ERRORS.VALUE);
	});
});

describe('Treasury bills', () => {
	const s = date(2008, 3, 31);
	const m = date(2008, 6, 1);

	it('match the documented examples', () => {
		expect(f('TBILLEQ', s, m, 0.0914)).toBeCloseTo(0.094151493565943, 13);
		expect(f('TBILLPRICE', s, m, 0.09)).toBeCloseTo(98.45, 12);
		expect(f('TBILLYIELD', s, m, 98.45)).toBeCloseTo(0.0914169629253426, 13);
	});

	it('round-trip price and yield', () => {
		const price = num('TBILLPRICE', s, m, 0.0525);
		expect(f('TBILLYIELD', s, m, price)).toBeCloseTo(((100 - price) / price) * (360 / 62), 14);
	});

	it('allow at most one year to maturity', () => {
		expect(typeof f('TBILLPRICE', date(2008, 1, 1), date(2009, 1, 1), 0.05)).toBe('number');
		expect(f('TBILLPRICE', date(2008, 1, 1), date(2009, 1, 2), 0.05)).toBe(ERRORS.NUM);
		expect(f('TBILLEQ', m, s, 0.05)).toBe(ERRORS.NUM);
		expect(f('TBILLEQ', s, s, 0.05)).toBe(ERRORS.NUM);
	});

	it('validate the rate or price', () => {
		expect(f('TBILLEQ', s, m, 0)).toBe(ERRORS.NUM);
		expect(f('TBILLPRICE', s, m, -0.05)).toBe(ERRORS.NUM);
		expect(f('TBILLPRICE', s, m, 6)).toBe(ERRORS.NUM);
		expect(f('TBILLYIELD', s, m, 0)).toBe(ERRORS.NUM);
		expect(f('TBILLYIELD', 'x', m, 98)).toBe(ERRORS.VALUE);
	});
});

describe('DOLLARDE / DOLLARFR', () => {
	it('match the documented examples', () => {
		expect(f('DOLLARDE', 1.02, 16)).toBeCloseTo(1.125, 14);
		expect(f('DOLLARDE', 1.1, 32)).toBeCloseTo(1.3125, 14);
		expect(f('DOLLARFR', 1.125, 16)).toBeCloseTo(1.02, 14);
		expect(f('DOLLARFR', 1.125, 32)).toBeCloseTo(1.04, 14);
	});

	it('handle signs, powers of ten and truncated fractions', () => {
		expect(f('DOLLARDE', -1.02, 16)).toBeCloseTo(-1.125, 14);
		expect(f('DOLLARFR', -1.125, 16)).toBeCloseTo(-1.02, 14);
		expect(f('DOLLARDE', 1.5, 10)).toBeCloseTo(1.5, 14);
		expect(f('DOLLARDE', 1.05, 100)).toBeCloseTo(1.05, 14);
		expect(f('DOLLARDE', 1.5, 1)).toBeCloseTo(1.5, 14);
		expect(f('DOLLARDE', 1.02, 16.9)).toBeCloseTo(1.125, 14);
		expect(f('DOLLARFR', f('DOLLARDE', 3.07, 8) as number, 8)).toBeCloseTo(3.07, 13);
	});

	it('validate the fraction', () => {
		expect(f('DOLLARDE', 1.02, -1)).toBe(ERRORS.NUM);
		expect(f('DOLLARDE', 1.02, 0.5)).toBe(ERRORS.DIV0);
		expect(f('DOLLARFR', 1.02, 0)).toBe(ERRORS.DIV0);
		expect(f('DOLLARFR', 'x', 2)).toBe(ERRORS.VALUE);
	});
});

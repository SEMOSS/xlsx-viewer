import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);
const num = (name: string, ...args: FunctionArg[]) => f(name, ...args) as number;

describe('SLN / SYD', () => {
	it('match the documented examples', () => {
		expect(f('SLN', 30000, 7500, 10)).toBe(2250);
		expect(f('SYD', 30000, 7500, 10, 1)).toBeCloseTo(4090.90909090909, 10);
		expect(f('SYD', 30000, 7500, 10, 10)).toBeCloseTo(409.090909090909, 10);
	});

	it('SYD depreciates down to the salvage value', () => {
		let total = 0;
		for (let per = 1; per <= 10; per++) {
			total += num('SYD', 30000, 7500, 10, per);
		}
		expect(total).toBeCloseTo(22500, 9);
	});

	it('validate arguments', () => {
		expect(f('SLN', 30000, 7500, 0)).toBe(ERRORS.DIV0);
		expect(f('SYD', 30000, 7500, 10, 0)).toBe(ERRORS.NUM);
		expect(f('SYD', 30000, 7500, 10, 11)).toBe(ERRORS.NUM);
		expect(f('SYD', 30000, 7500, 0, 1)).toBe(ERRORS.NUM);
		expect(f('SLN', 'x', 1, 1)).toBe(ERRORS.VALUE);
	});
});

describe('DB', () => {
	it.each([
		[1, 186083.33],
		[2, 259639.42],
		[3, 176814.44],
		[4, 120410.64],
		[5, 81999.64],
		[6, 55841.76],
		[7, 15845.1],
	])('documented example, period %d', (period, value) => {
		expect(f('DB', 1000000, 100000, 6, period, 7)).toBeCloseTo(value, 2);
	});

	it('defaults to a full first year and rounds the rate to three decimals', () => {
		expect(f('DB', 1000000, 100000, 6, 1)).toBeCloseTo(319000, 8);
		expect(f('DB', 1000000, 100000, 6, 7)).toBe(0);
		expect(f('DB', 10000, 0, 5, 1)).toBe(10000);
	});

	it('validates arguments', () => {
		expect(f('DB', 1000000, 100000, 6, 8, 7)).toBe(ERRORS.NUM);
		expect(f('DB', 1000000, 100000, 6, 1, 13)).toBe(ERRORS.NUM);
		expect(f('DB', 1000000, 100000, 6, 1, 0)).toBe(ERRORS.NUM);
		expect(f('DB', 0, 0, 6, 1)).toBe(ERRORS.NUM);
		expect(f('DB', 1000, 2000, 6, 1)).toBe(ERRORS.NUM);
		expect(f('DB', 1000, -1, 6, 1)).toBe(ERRORS.NUM);
		expect(f('DB', 1000, 100, 0, 1)).toBe(ERRORS.NUM);
		expect(f('DB', 1000, 100, 6, 0)).toBe(ERRORS.NUM);
	});
});

describe('DDB', () => {
	it('matches the documented examples', () => {
		expect(f('DDB', 2400, 300, 10 * 365, 1)).toBeCloseTo(1.31506849315068, 12);
		expect(f('DDB', 2400, 300, 10 * 12, 1, 2)).toBeCloseTo(40, 12);
		expect(f('DDB', 2400, 300, 10, 1, 2)).toBeCloseTo(480, 12);
		expect(f('DDB', 2400, 300, 10, 2, 1.5)).toBeCloseTo(306, 10);
		expect(f('DDB', 2400, 300, 10, 10)).toBeCloseTo(22.1225472, 7);
	});

	it('never depreciates below the salvage value', () => {
		expect(f('DDB', 2400, 300, 2, 1)).toBe(2100);
		expect(f('DDB', 2400, 300, 2, 2)).toBe(0);
		let total = 0;
		for (let per = 1; per <= 10; per++) {
			total += num('DDB', 2400, 300, 10, per);
		}
		expect(total).toBeCloseTo(2100, 9);
	});

	it('validates arguments', () => {
		expect(f('DDB', 2400, 300, 10, 11)).toBe(ERRORS.NUM);
		expect(f('DDB', 2400, 300, 10, 0.5)).toBe(ERRORS.NUM);
		expect(f('DDB', 2400, 300, 10, 1, 0)).toBe(ERRORS.NUM);
		expect(f('DDB', 2400, 3000, 10, 1)).toBe(ERRORS.NUM);
		expect(f('DDB', -1, 0, 10, 1)).toBe(ERRORS.NUM);
	});
});

describe('VDB', () => {
	it('matches the documented examples', () => {
		expect(f('VDB', 2400, 300, 10 * 365, 0, 1)).toBeCloseTo(1.31506849315068, 12);
		expect(f('VDB', 2400, 300, 10 * 12, 0, 1)).toBeCloseTo(40, 12);
		expect(f('VDB', 2400, 300, 10, 0, 1)).toBeCloseTo(480, 12);
		expect(f('VDB', 2400, 300, 10 * 12, 6, 18)).toBeCloseTo(396.306053264752, 9);
		expect(f('VDB', 2400, 300, 10 * 12, 6, 18, 1.5)).toBeCloseTo(311.808936658233, 9);
		expect(f('VDB', 2400, 300, 10, 0, 0.875, 1.5)).toBeCloseTo(315, 10);
	});

	it('depreciates exactly to the salvage value over the whole life', () => {
		expect(f('VDB', 2400, 300, 10, 0, 10)).toBeCloseTo(2100, 9);
		expect(f('VDB', 2400, 300, 10, 0, 10, 2, true)).toBeCloseTo(2100, 9);
		const first = num('VDB', 2400, 300, 10, 0, 5);
		const second = num('VDB', 2400, 300, 10, 5, 10);
		expect(first + second).toBeCloseTo(2100, 9);
	});

	it('switches to straight line only when allowed', () => {
		const late = num('VDB', 1000, 0, 5, 3, 4);
		const lateNoSwitch = num('VDB', 1000, 0, 5, 3, 4, 2, true);
		expect(late).toBeCloseTo(108, 9);
		expect(lateNoSwitch).toBeCloseTo(86.4, 9);
		expect(f('VDB', 2400, 300, 10, 0, 1, 2, true)).toBeCloseTo(480, 12);
	});

	it('validates arguments', () => {
		expect(f('VDB', 2400, 300, 10, 2, 1)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 300, 10, -1, 1)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 300, 10, 0, 11)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 300, 10, 0, 1, 0)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 3000, 10, 0, 1)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 300, 0, 0, 0)).toBe(ERRORS.NUM);
		expect(f('VDB', 2400, 300, 10, 0, 1, 2, 'maybe')).toBe(ERRORS.VALUE);
	});
});

describe('AMORDEGRC / AMORLINC', () => {
	const purchased = 39679; // 2008-08-19
	const firstPeriod = 39813; // 2008-12-31

	it('match the documented examples', () => {
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, 1, 0.15, 1)).toBe(776);
		expect(f('AMORLINC', 2400, purchased, firstPeriod, 300, 1, 0.15, 1)).toBe(360);
	});

	it.each([
		[0, 330],
		[1, 776],
		[2, 485],
		[3, 303],
		[4, 190],
		[5, 158],
		[6, 0],
	])('AMORDEGRC period %d = %d', (period, value) => {
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, period, 0.15, 1)).toBe(value);
	});

	it.each([
		[0, 131.803278688525],
		[1, 360],
		[5, 360],
		[6, 168.196721311475],
		[7, 0],
	])('AMORLINC period %d = %d', (period, value) => {
		expect(f('AMORLINC', 2400, purchased, firstPeriod, 300, period, 0.15, 1)).toBeCloseTo(value, 9);
	});

	it('validate arguments', () => {
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, 1, 0.15, 2)).toBe(ERRORS.NUM);
		expect(f('AMORLINC', 2400, purchased, firstPeriod, 300, 1, 0.15, 5)).toBe(ERRORS.NUM);
		expect(f('AMORLINC', 2400, purchased, firstPeriod, 300, 1, 0)).toBe(ERRORS.NUM);
		expect(f('AMORLINC', 2400, firstPeriod, purchased, 300, 1, 0.15)).toBe(ERRORS.NUM);
		expect(f('AMORLINC', 2400, purchased, firstPeriod, 3000, 1, 0.15)).toBe(ERRORS.NUM);
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, 1, 0.22)).toBe(ERRORS.NUM);
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, 1, 0.6)).toBe(ERRORS.NUM);
		expect(f('AMORDEGRC', 2400, purchased, firstPeriod, 300, 1, 0.5)).toBe(980);
		expect(f('AMORDEGRC', 2400, 'x', firstPeriod, 300, 1, 0.15)).toBe(ERRORS.VALUE);
	});
});

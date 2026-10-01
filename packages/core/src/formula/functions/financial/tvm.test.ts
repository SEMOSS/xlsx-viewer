import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);
const num = (name: string, ...args: FunctionArg[]) => f(name, ...args) as number;

describe('PMT / FV / PV', () => {
	it('matches the documented examples', () => {
		expect(f('PMT', 0.05 / 12, 360, 200000)).toBeCloseTo(-1073.64324602798, 8);
		expect(f('PMT', 0.08 / 12, 10, 10000)).toBeCloseTo(-1037.032089, 6);
		expect(f('PMT', 0.08 / 12, 10, 10000, 0, 1)).toBeCloseTo(-1030.164327, 6);
		expect(f('PMT', 0.06 / 12, 18 * 12, 0, 50000)).toBeCloseTo(-129.0811609, 6);
		expect(f('FV', 0.06 / 12, 10, -200, -500, 1)).toBeCloseTo(2581.403374, 6);
		expect(f('FV', 0.12 / 12, 12, -1000)).toBeCloseTo(12682.50301, 5);
		expect(f('FV', 0.11 / 12, 35, -2000, null, 1)).toBeCloseTo(82846.24637, 5);
		expect(f('FV', 0.06 / 12, 12, -100, -1000, 1)).toBeCloseTo(2301.40183, 5);
		expect(f('PV', 0.08 / 12, 12 * 20, 500, null, 0)).toBeCloseTo(-59777.14585, 5);
	});

	it('handles a zero rate', () => {
		expect(f('PMT', 0, 10, 1000)).toBe(-100);
		expect(f('PMT', 0, 10, 1000, 500)).toBe(-150);
		expect(f('FV', 0, 10, -100)).toBe(1000);
		expect(f('FV', 0, 10, -100, -50)).toBe(1050);
		expect(f('PV', 0, 10, -100)).toBe(1000);
	});

	it('treats any non-zero type as payment in advance', () => {
		expect(f('PMT', 0.01, 12, 1000, 0, 2)).toBe(f('PMT', 0.01, 12, 1000, 0, 1));
		expect(f('PV', 0.01, 12, -100, 0, true)).toBe(f('PV', 0.01, 12, -100, 0, 1));
	});

	it('round-trips between PV, FV and PMT', () => {
		const pmt = num('PMT', 0.004, 120, 25000, -5000, 1);
		expect(f('PV', 0.004, 120, pmt, -5000, 1)).toBeCloseTo(25000, 8);
		expect(f('FV', 0.004, 120, pmt, 25000, 1)).toBeCloseTo(-5000, 8);
	});

	it('validates arguments', () => {
		expect(f('PMT', 0.1, 0, 1000)).toBe(ERRORS.NUM);
		expect(f('PMT', 0, 0, 1000)).toBe(ERRORS.NUM);
		expect(f('PMT', 'x', 10, 1000)).toBe(ERRORS.VALUE);
		expect(f('PV', ERRORS.NA, 10, 1000)).toBe(ERRORS.NA);
		expect(f('FV', -2, 0.5, 100)).toBe(ERRORS.NUM);
		expect(f('PMT', 0.1)).toBe(ERRORS.VALUE);
	});

	it('lifts over arrays', () => {
		expect(f('PMT', 0, arr([10, 20]), 1000)).toEqual(arr([-100, -50]));
	});
});

describe('NPER', () => {
	it('matches the documented examples', () => {
		expect(f('NPER', 0.12 / 12, -100, -1000, 10000, 1)).toBeCloseTo(59.6738657, 7);
		expect(f('NPER', 0.12 / 12, -100, -1000, 10000)).toBeCloseTo(60.0821229, 7);
		expect(f('NPER', 0.12 / 12, -100, -1000)).toBeCloseTo(-9.57859404, 7);
	});

	it('handles a zero rate and impossible terms', () => {
		expect(f('NPER', 0, -100, 1000)).toBe(10);
		expect(f('NPER', 0, 0, 1000)).toBe(ERRORS.NUM);
		expect(f('NPER', 0.1, -100, 2000)).toBe(ERRORS.NUM);
		expect(f('NPER', 0.1, 100, 1000)).toBeCloseTo(Math.log(0.5) / Math.log(1.1), 13);
	});
});

describe('IPMT / PPMT', () => {
	it('matches the documented examples', () => {
		expect(f('IPMT', 0.1 / 12, 1, 36, 8000)).toBeCloseTo(-66.66666667, 7);
		expect(f('IPMT', 0.1, 3, 3, 8000)).toBeCloseTo(-292.4471299, 7);
		expect(f('PPMT', 0.1 / 12, 1, 24, 2000)).toBeCloseTo(-75.62318601, 7);
		expect(f('PPMT', 0.08, 10, 10, 200000)).toBeCloseTo(-27598.05346, 5);
	});

	it('splits every payment into interest and principal', () => {
		for (const type of [0, 1]) {
			const pmt = num('PMT', 0.01, 24, 5000, 0, type);
			let principal = 0;
			for (let per = 1; per <= 24; per++) {
				const i = num('IPMT', 0.01, per, 24, 5000, 0, type);
				const p = num('PPMT', 0.01, per, 24, 5000, 0, type);
				expect(i + p).toBeCloseTo(pmt, 9);
				principal += p;
			}
			expect(principal).toBeCloseTo(-5000, 8);
		}
	});

	it('has no interest in the first period when paying in advance', () => {
		expect(f('IPMT', 0.01, 1, 12, 1000, 0, 1)).toBe(0);
		expect(f('PPMT', 0.01, 1, 12, 1000, 0, 1)).toBe(f('PMT', 0.01, 12, 1000, 0, 1));
	});

	it('requires 1 <= per <= nper', () => {
		expect(f('IPMT', 0.1, 0, 3, 8000)).toBe(ERRORS.NUM);
		expect(f('IPMT', 0.1, 4, 3, 8000)).toBe(ERRORS.NUM);
		expect(f('PPMT', 0.1, -1, 3, 8000)).toBe(ERRORS.NUM);
		expect(f('IPMT', 0, 2, 3, 8000)).toBe(0);
	});
});

describe('ISPMT', () => {
	it('matches the documented examples', () => {
		expect(f('ISPMT', 0.1 / 12, 1, 36, 8000000)).toBeCloseTo(-64814.81481, 5);
		expect(f('ISPMT', 0.1, 1, 3, 8000000)).toBeCloseTo(-533333.3333, 4);
		expect(f('ISPMT', 0.1, 0, 4, 4000)).toBe(-400);
	});

	it('rejects zero periods', () => {
		expect(f('ISPMT', 0.1, 1, 0, 4000)).toBe(ERRORS.DIV0);
	});
});

describe('CUMIPMT / CUMPRINC', () => {
	const rate = 0.09 / 12;

	it('matches the documented examples', () => {
		expect(f('CUMIPMT', rate, 360, 125000, 13, 24, 0)).toBeCloseTo(-11135.23213, 5);
		expect(f('CUMIPMT', rate, 360, 125000, 1, 1, 0)).toBeCloseTo(-937.5, 9);
		expect(f('CUMPRINC', rate, 360, 125000, 13, 24, 0)).toBeCloseTo(-934.1071234, 7);
		expect(f('CUMPRINC', rate, 360, 125000, 1, 1, 0)).toBeCloseTo(-68.27827118, 8);
	});

	it('adds up the whole loan', () => {
		const pmt = num('PMT', rate, 360, 125000);
		const interest = num('CUMIPMT', rate, 360, 125000, 1, 360, 0);
		const principal = num('CUMPRINC', rate, 360, 125000, 1, 360, 0);
		expect(principal).toBeCloseTo(-125000, 6);
		expect(interest + principal).toBeCloseTo(pmt * 360, 6);
		const advance = num('CUMPRINC', rate, 360, 125000, 1, 360, 1);
		expect(advance).toBeCloseTo(-125000, 6);
		expect(f('CUMIPMT', rate, 360, 125000, 1, 1, 1)).toBe(0);
	});

	it('truncates periods and validates the terms', () => {
		expect(f('CUMIPMT', rate, 360.9, 125000, 13.7, 24.2, 0)).toBe(
			f('CUMIPMT', rate, 360, 125000, 13, 24, 0),
		);
		expect(f('CUMIPMT', 0, 360, 125000, 1, 2, 0)).toBe(ERRORS.NUM);
		expect(f('CUMIPMT', rate, 0, 125000, 1, 2, 0)).toBe(ERRORS.NUM);
		expect(f('CUMIPMT', rate, 360, -1, 1, 2, 0)).toBe(ERRORS.NUM);
		expect(f('CUMIPMT', rate, 360, 125000, 0, 2, 0)).toBe(ERRORS.NUM);
		expect(f('CUMIPMT', rate, 360, 125000, 3, 2, 0)).toBe(ERRORS.NUM);
		expect(f('CUMPRINC', rate, 360, 125000, 1, 361, 0)).toBe(ERRORS.NUM);
		expect(f('CUMPRINC', rate, 360, 125000, 1, 2, 2)).toBe(ERRORS.NUM);
		expect(f('CUMPRINC', rate, 360, 125000, 1, 2)).toBe(ERRORS.VALUE);
	});
});

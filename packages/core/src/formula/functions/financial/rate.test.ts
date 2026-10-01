import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 0.1, A3: 0.2, B1: 0.1, B2: 'x' });
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);
const num = (name: string, ...args: FunctionArg[]) => f(name, ...args) as number;

describe('RATE', () => {
	it('matches the documented example', () => {
		expect(f('RATE', 48, -200, 8000)).toBeCloseTo(0.00770147248820165, 15);
		expect(num('RATE', 48, -200, 8000) * 12).toBeCloseTo(0.0924176698584198, 13);
	});

	it('inverts PV for both payment types', () => {
		for (const type of [0, 1]) {
			for (const r of [0.001, 0.0125, 0.05, 0.3]) {
				const pv = num('PV', r, 36, -250, 1000, type);
				// These cash flows change sign twice; the guess selects the root near 0.3.
				const guess = r > 0.1 ? 0.35 : 0.1;
				expect(f('RATE', 36, -250, pv, 1000, type, guess)).toBeCloseTo(r, 12);
			}
		}
		expect(f('RATE', 36, -250, num('PV', 0.3, 36, -250, 1000), 1000)).toBeCloseTo(
			-0.249985423779533,
			12,
		);
	});

	it('handles fractional periods, guesses and a zero rate', () => {
		const pv = num('PV', 0.02, 10.5, -100);
		expect(f('RATE', 10.5, -100, pv)).toBeCloseTo(0.02, 12);
		expect(f('RATE', 48, -200, 8000, 0, 0, 0.05)).toBeCloseTo(0.00770147248820165, 15);
		expect(Math.abs(num('RATE', 10, -100, 1000))).toBeLessThan(1e-12);
		expect(f('RATE', 10, 0, -1000, 2000)).toBeCloseTo(Math.pow(2, 0.1) - 1, 13);
	});

	it('returns #NUM! when Newton does not converge within 20 iterations', () => {
		expect(f('RATE', 48, -200, 8000, 0, 0, 0.5)).toBe(ERRORS.NUM);
		expect(f('RATE', 10, 100, 1000)).toBe(ERRORS.NUM);
		expect(f('RATE', 0, -100, 1000)).toBe(ERRORS.NUM);
		expect(f('RATE', -5, -100, 1000)).toBe(ERRORS.NUM);
		expect(f('RATE', 'x', -100, 1000)).toBe(ERRORS.VALUE);
	});
});

describe('EFFECT / NOMINAL', () => {
	it('matches the documented examples', () => {
		expect(f('EFFECT', 0.0525, 4)).toBeCloseTo(0.0535426673707582, 14);
		expect(f('NOMINAL', 0.053543, 4)).toBeCloseTo(0.052500319868356, 14);
	});

	it('are inverses and truncate npery', () => {
		expect(f('NOMINAL', num('EFFECT', 0.07, 12), 12)).toBeCloseTo(0.07, 14);
		expect(f('EFFECT', 0.05, 4.9)).toBe(f('EFFECT', 0.05, 4));
		expect(f('EFFECT', 0.1, 1)).toBeCloseTo(0.1, 15);
	});

	it('validate arguments', () => {
		expect(f('EFFECT', 0, 4)).toBe(ERRORS.NUM);
		expect(f('EFFECT', 0.05, 0.5)).toBe(ERRORS.NUM);
		expect(f('NOMINAL', -0.01, 4)).toBe(ERRORS.NUM);
		expect(f('NOMINAL', 0.05, 0)).toBe(ERRORS.NUM);
		expect(f('EFFECT', 'a', 4)).toBe(ERRORS.VALUE);
	});
});

describe('PDURATION / RRI', () => {
	it('match the documented examples', () => {
		expect(f('PDURATION', 0.025, 2000, 2200)).toBeCloseTo(3.85986616262264, 13);
		expect(f('PDURATION', 0.025 / 12, 1000, 1200)).toBeCloseTo(87.6054764193714, 10);
		expect(f('RRI', 96, 10000, 11000)).toBeCloseTo(0.000993307376291, 14);
	});

	it('validate arguments', () => {
		expect(f('PDURATION', 0, 2000, 2200)).toBe(ERRORS.NUM);
		expect(f('PDURATION', 0.025, -2000, 2200)).toBe(ERRORS.NUM);
		expect(f('PDURATION', 0.025, 2000, 0)).toBe(ERRORS.NUM);
		expect(f('RRI', 0, 10000, 11000)).toBe(ERRORS.NUM);
		expect(f('RRI', 10, 0, 11000)).toBe(ERRORS.NUM);
		expect(f('RRI', 10, -100, 200)).toBe(ERRORS.NUM);
		expect(f('RRI', 2, -100, -400)).toBeCloseTo(1, 15);
	});

	it('carry the _xlfn. file prefix', () => {
		expect(fn(FINANCIAL_FUNCTIONS, 'PDURATION').filePrefix).toBe('_xlfn.');
		expect(fn(FINANCIAL_FUNCTIONS, 'RRI').filePrefix).toBe('_xlfn.');
		expect(fn(FINANCIAL_FUNCTIONS, 'RATE').filePrefix).toBeUndefined();
	});
});

describe('FVSCHEDULE', () => {
	it('matches the documented example', () => {
		expect(f('FVSCHEDULE', 1, arr([0.09, 0.11, 0.1]))).toBeCloseTo(1.33089, 14);
	});

	it('treats blanks as zero and rejects other values', () => {
		expect(f('FVSCHEDULE', 100, ctx.ref('A1:A3'))).toBeCloseTo(132, 12);
		expect(f('FVSCHEDULE', 100, ctx.ref('B1:B2'))).toBe(ERRORS.VALUE);
		expect(f('FVSCHEDULE', 100, arr([0.1, true]))).toBe(ERRORS.VALUE);
		expect(f('FVSCHEDULE', 100, arr([0.1, ERRORS.DIV0]))).toBe(ERRORS.DIV0);
		expect(f('FVSCHEDULE', 100, 0.5)).toBe(150);
		expect(f('FVSCHEDULE', 100, '50%')).toBe(150);
		expect(f('FVSCHEDULE', 100, 'abc')).toBe(ERRORS.VALUE);
	});
});

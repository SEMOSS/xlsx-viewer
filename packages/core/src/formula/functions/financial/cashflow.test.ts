import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { FINANCIAL_FUNCTIONS } from './index';

const ctx = new MockContext({
	A1: -10000,
	A2: 3000,
	A3: 'note',
	A4: 4200,
	A5: true,
	A6: 6800,
	B1: -100,
	B2: ERRORS.DIV0,
});
const f = (name: string, ...args: FunctionArg[]) => call(fn(FINANCIAL_FUNCTIONS, name), args, ctx);

const flows = arr([-10000, 2750, 4250, 3250, 2750]);
const dates = arr([39448, 39508, 39751, 39859, 39904]);

describe('NPV', () => {
	it('matches the documented examples', () => {
		expect(f('NPV', 0.1, -10000, 3000, 4200, 6800)).toBeCloseTo(1188.44341233522, 9);
		const npv = f('NPV', 0.08, 8000, 9200, 10000, 12000, 14500) as number;
		expect(npv - 40000).toBeCloseTo(1922.06, 2);
		const withLoss = f('NPV', 0.08, arr([8000, 9200, 10000, 12000, 14500]), -9000) as number;
		expect(withLoss - 40000).toBeCloseTo(-3749.47, 2);
	});

	it('ignores text and logicals in references, coerces direct arguments', () => {
		expect(f('NPV', 0.1, ctx.ref('A1:A6'))).toBeCloseTo(1188.44341233522, 9);
		expect(f('NPV', 0.1, '100', true)).toBeCloseTo(100 / 1.1 + 1 / 1.21, 12);
		expect(f('NPV', 0, 1, 2, 3)).toBe(6);
	});

	it('propagates errors and rejects a rate of -100%', () => {
		expect(f('NPV', 0.1, ctx.ref('B1:B2'))).toBe(ERRORS.DIV0);
		expect(f('NPV', 0.1, 'abc')).toBe(ERRORS.VALUE);
		expect(f('NPV', -1, 100)).toBe(ERRORS.DIV0);
	});
});

describe('IRR', () => {
	it('matches the documented examples', () => {
		expect(f('IRR', arr([-70000, 12000, 15000, 18000, 21000]))).toBeCloseTo(-0.0212448482734, 12);
		expect(f('IRR', arr([-70000, 12000, 15000, 18000, 21000, 26000]))).toBeCloseTo(
			0.0866309480365,
			12,
		);
		expect(f('IRR', arr([-70000, 12000, 15000]), -0.1)).toBeCloseTo(-0.443506941334741, 12);
	});

	it('zeroes the NPV at the result', () => {
		const values = [-1000, 300, 400, 500, 100];
		const r = f('IRR', arr(values)) as number;
		const npv = values.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0);
		expect(Math.abs(npv)).toBeLessThan(1e-9);
	});

	it('falls back to a bracketing search when Newton fails', () => {
		expect(f('IRR', arr([-1, 100]), 1000)).toBeCloseTo(99, 10);
		expect(f('IRR', arr([-100, 0, 0, 0, 1e6]))).toBeCloseTo(9, 10);
	});

	it('ignores text and logicals in references', () => {
		expect(f('IRR', ctx.ref('A1:A6'))).toBeCloseTo(
			f('IRR', arr([-10000, 3000, 4200, 6800])) as number,
			14,
		);
	});

	it('needs both positive and negative cash flows', () => {
		expect(f('IRR', arr([100, 200]))).toBe(ERRORS.NUM);
		expect(f('IRR', arr([-100, -200]))).toBe(ERRORS.NUM);
		expect(f('IRR', ctx.ref('B1:B2'))).toBe(ERRORS.DIV0);
	});
});

describe('MIRR', () => {
	const values = arr([-120000, 39000, 30000, 21000, 37000, 46000]);

	it('matches the documented examples', () => {
		expect(f('MIRR', values, 0.1, 0.12)).toBeCloseTo(0.126094130365905, 12);
		expect(f('MIRR', arr([-120000, 39000, 30000, 21000]), 0.1, 0.12)).toBeCloseTo(-0.048044655, 9);
		expect(f('MIRR', values, 0.1, 0.14)).toBeCloseTo(0.134759110828315, 12);
	});

	it('needs both positive and negative cash flows', () => {
		expect(f('MIRR', arr([1, 2, 3]), 0.1, 0.1)).toBe(ERRORS.DIV0);
		expect(f('MIRR', values, -1, 0.1)).toBe(ERRORS.DIV0);
		expect(f('MIRR', ctx.ref('B1:B2'), 0.1, 0.1)).toBe(ERRORS.DIV0);
	});
});

describe('XNPV / XIRR', () => {
	it('match the documented examples', () => {
		expect(f('XNPV', 0.09, flows, dates)).toBeCloseTo(2086.64760203154, 8);
		expect(f('XIRR', flows, dates)).toBeCloseTo(0.373362533518832, 9);
	});

	it('accept dates in any order after the first, as text, and a guess', () => {
		const shuffled = arr([-10000, 3250, 2750, 4250, 2750]);
		const shuffledDates = arr([39448, 39859, 39508, 39751, 39904]);
		expect(f('XNPV', 0.09, shuffled, shuffledDates)).toBeCloseTo(2086.64760203154, 8);
		expect(f('XIRR', flows, dates, 0.5)).toBeCloseTo(0.373362533518832, 9);
		expect(f('XNPV', 0.09, arr([-100, 110]), arr(['1/1/2024', '1/1/2025']))).toBeCloseTo(
			-100 + 110 / Math.pow(1.09, 366 / 365),
			12,
		);
		expect(f('XIRR', arr([-100, 110]), arr([45292, 45657]))).toBeCloseTo(0.1, 12);
	});

	it('skip blank pairs', () => {
		expect(
			f(
				'XNPV',
				0.09,
				arr([-10000, 2750, null], [4250, 3250, 2750]),
				arr([39448, 39508, null], [39751, 39859, 39904]),
			),
		).toBeCloseTo(2086.64760203154, 8);
	});

	it('validate the cash flows and dates', () => {
		expect(f('XNPV', 0.09, arr([-100, 100]), arr([45292]))).toBe(ERRORS.NUM);
		expect(f('XNPV', 0.09, arr([-100, 100]), arr([45292, 45000]))).toBe(ERRORS.NUM);
		expect(f('XNPV', 0.09, arr([-100, 'x']), arr([45292, 45300]))).toBe(ERRORS.VALUE);
		expect(f('XNPV', 0.09, arr([-100, 100]), arr([45292, 'soon']))).toBe(ERRORS.VALUE);
		expect(f('XNPV', 0.09, arr([-100, 100]), arr([45292, -1]))).toBe(ERRORS.NUM);
		expect(f('XNPV', -1, arr([-100, 100]), arr([45292, 45300]))).toBe(ERRORS.NUM);
		expect(f('XNPV', 0.09, arr([-100, ERRORS.NA]), arr([45292, 45300]))).toBe(ERRORS.NA);
		expect(f('XIRR', arr([100, 100]), arr([45292, 45300]))).toBe(ERRORS.NUM);
		expect(f('XIRR', flows, dates, -1)).toBe(ERRORS.NUM);
	});
});

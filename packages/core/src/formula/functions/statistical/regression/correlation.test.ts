import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, stat } from '../test-utils/expect';

const Y7 = arr([2, 3, 9, 1, 8, 7, 5]);
const X7 = arr([6, 5, 11, 7, 5, 4, 4]);

// Expected values are the exact rational results of the Excel documentation examples.
describe('correlation and covariance', () => {
	it('matches Excel', () => {
		expectClose(
			stat('CORREL', [arr([3, 2, 4, 5, 6]), arr([9, 7, 12, 15, 17])]),
			0.9970544855015816,
			1e-15,
		);
		expectClose(
			stat('PEARSON', [arr([9, 7, 5, 3, 1]), arr([10, 6, 1, 5, 3])]),
			0.6993786061802353,
			1e-15,
		);
		expectClose(stat('RSQ', [Y7, X7]), 0.05795019157088123, 1e-14);
		expectClose(stat('COVAR', [arr([3, 2, 4, 5, 6]), arr([9, 7, 12, 15, 17])]), 5.2, 1e-15);
		expectClose(stat('COVARIANCE.P', [arr([3, 2, 4, 5, 6]), arr([9, 7, 12, 15, 17])]), 5.2, 1e-15);
		expectClose(stat('COVARIANCE.S', [arr([2, 4, 8]), arr([5, 11, 12])]), 9.666666666666666, 1e-15);
	});

	it('drops pairs where either value is not a number', () => {
		const ctx = new MockContext({
			A1: 3,
			A2: 'x',
			A3: 4,
			A4: 5,
			A5: 6,
			A6: 2,
			B1: 9,
			B2: 1,
			B3: 12,
			B4: 15,
			B5: 17,
			B6: true,
		});
		// (x, 1) and (2, TRUE) are ignored, leaving (3, 9), (4, 12), (5, 15), (6, 17)
		expectClose(
			stat('CORREL', [ctx.ref('A1:A6'), ctx.ref('B1:B6')], ctx),
			stat('CORREL', [arr([3, 4, 5, 6]), arr([9, 12, 15, 17])]) as number,
			1e-15,
		);
	});

	it('pairs arrays of different shapes with the same size in row-major order', () => {
		expectClose(
			stat('COVAR', [arr([3, 2], [4, 5]), arr([9, 7, 12, 15])]),
			stat('COVAR', [arr([3, 2, 4, 5]), arr([9, 7, 12, 15])]) as number,
			1e-15,
		);
	});

	it('validates the data', () => {
		expect(stat('CORREL', [arr([1, 2, 3]), arr([1, 2])])).toBe(ERRORS.NA);
		expect(stat('CORREL', [arr(['a']), arr(['b'])])).toBe(ERRORS.DIV0);
		expect(stat('CORREL', [arr([1, 1, 1]), arr([1, 2, 3])])).toBe(ERRORS.DIV0);
		expect(stat('PEARSON', [arr(['a']), arr(['b'])])).toBe(ERRORS.NA);
		expect(stat('RSQ', [arr([1]), arr([2])])).toBe(ERRORS.DIV0);
		expect(stat('COVAR', [arr(['a']), arr([1])])).toBe(ERRORS.DIV0);
		expect(stat('COVARIANCE.S', [arr([1]), arr([2])])).toBe(ERRORS.DIV0);
		expect(stat('CORREL', [arr([1, ERRORS.NUM]), arr([1, 2])])).toBe(ERRORS.NUM);
	});
});

describe('simple linear regression', () => {
	it('matches Excel', () => {
		expectClose(stat('SLOPE', [Y7, X7]), 0.3055555555555556, 1e-14);
		expectClose(
			stat('INTERCEPT', [arr([2, 3, 9, 1, 8]), arr([6, 5, 11, 7, 5])]),
			0.04838709677419355,
			1e-12,
		);
		expectClose(stat('STEYX', [Y7, X7]), 3.3057189502100415, 1e-14);
		const y = arr([6, 7, 9, 15, 21]);
		const x = arr([20, 28, 31, 38, 40]);
		expectClose(stat('FORECAST', [30, y, x]), 10.607253086419753, 1e-14);
		expectClose(stat('FORECAST.LINEAR', [30, y, x]), 10.607253086419753, 1e-14);
	});

	it('is exact for collinear data', () => {
		const x = arr([1, 2, 3, 4]);
		const y = arr([3, 5, 7, 9]);
		expect(stat('SLOPE', [y, x])).toBe(2);
		expect(stat('INTERCEPT', [y, x])).toBe(1);
		expect(stat('STEYX', [y, x])).toBe(0);
		expect(stat('RSQ', [y, x])).toBe(1);
		expect(stat('FORECAST', [10, y, x])).toBe(21);
	});

	it('validates the data', () => {
		expect(stat('SLOPE', [arr([1, 2]), arr([1, 2, 3])])).toBe(ERRORS.NA);
		expect(stat('SLOPE', [arr(['a']), arr(['b'])])).toBe(ERRORS.NA);
		expect(stat('SLOPE', [arr([1, 2]), arr([3, 3])])).toBe(ERRORS.DIV0);
		expect(stat('INTERCEPT', [arr([1, 2]), arr([3, 3])])).toBe(ERRORS.DIV0);
		expect(stat('STEYX', [arr([1, 2]), arr([3, 4])])).toBe(ERRORS.DIV0);
		expect(stat('FORECAST', ['x', arr([1, 2]), arr([3, 4])])).toBe(ERRORS.VALUE);
		expect(stat('FORECAST', [1, arr([1, 2]), arr([3, 3])])).toBe(ERRORS.DIV0);
	});
});

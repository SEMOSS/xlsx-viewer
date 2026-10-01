import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, stat } from '../test-utils/expect';

describe('F.DIST family', () => {
	it('matches Excel', () => {
		expectClose(stat('F.DIST', [15.2069, 6, 4, true]), 0.9900000430027627, 1e-14);
		expectClose(stat('F.DIST', [15.2069, 6, 4, false]), 0.0012237917087831731, 1e-14);
		expectClose(stat('F.DIST.RT', [15.2068649, 6, 4]), 0.009999999952464601, 1e-13);
		expectClose(stat('FDIST', [15.2068649, 6, 4]), 0.009999999952464601, 1e-13);
	});

	// [x, d1, d2, P(F <= x), P(F > x), density] from 60-digit references
	it.each([
		[1, 1, 1, 0.5, 0.5, 0.15915494309189535],
		[0.5, 5, 10, 0.22997511934989837, 0.7700248806501017, 0.6876070027706234],
		[2, 10, 20, 0.91021728515625, 0.08978271484375, 0.152740478515625],
		[100, 3, 7, 0.9999958695778987, 4.130422101383355e-6, 1.4090397251052762e-7],
		[0.01, 4, 9, 0.00023979488558792676, 0.9997602051144121, 0.047499788986201025],
		[1.5, 50, 60, 0.9335517053407213, 0.06644829465927873, 0.316733573943473],
	])('F.DIST(%d, %d, %d)', (x, d1, d2, cdf, sf, pdf) => {
		expectClose(stat('F.DIST', [x, d1, d2, true]), cdf, 1e-13);
		expectClose(stat('F.DIST', [x, d1, d2, false]), pdf, 1e-13);
		expectClose(stat('F.DIST.RT', [x, d1, d2]), sf, 1e-13);
	});

	it('truncates degrees of freedom and validates', () => {
		expect(stat('F.DIST', [1, 2.5, 3.9, true])).toBe(stat('F.DIST', [1, 2, 3, true]));
		expect(stat('F.DIST', [0, 2, 3, false])).toBe(1);
		expect(stat('F.DIST', [-1, 2, 3, true])).toBe(ERRORS.NUM);
		expect(stat('F.DIST', [1, 0, 3, true])).toBe(ERRORS.NUM);
		expect(stat('F.DIST.RT', [1, 2, 0.5])).toBe(ERRORS.NUM);
		expect(stat('FDIST', [1, 1e10, 3])).toBe(ERRORS.NUM);
		expect(stat('FDIST', [-2, 1, 3])).toBe(ERRORS.NUM);
	});
});

describe('F.INV family', () => {
	it('matches Excel', () => {
		expectClose(stat('F.INV', [0.01, 6, 4]), 0.10930991412457854, 1e-14);
		expectClose(stat('F.INV.RT', [0.01, 6, 4]), 15.20686486115753, 1e-14);
		expectClose(stat('FINV', [0.01, 6, 4]), 15.20686486115753, 1e-14);
	});

	it.each([
		['F.INV.RT', 0.05, 1, 1, 161.44763879758847],
		['F.INV.RT', 0.05, 5, 10, 3.325834530413012],
		['FINV', 0.001, 20, 3, 126.41777396613219],
		['F.INV', 0.025, 7, 12, 0.21432415255430798],
		['F.INV', 0.9, 2, 2, 9],
	])('%s(%d, %d, %d)', (name, p, d1, d2, x) => {
		expectClose(stat(name, [p, d1, d2]), x, 1e-14);
	});

	it('handles the end points and validates', () => {
		expect(stat('F.INV', [0, 2, 3])).toBe(0);
		expect(stat('F.INV.RT', [1, 2, 3])).toBe(0);
		expect(stat('F.INV', [1, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('F.INV.RT', [0, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('F.INV', [-0.5, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('FINV', [0.5, 0, 3])).toBe(ERRORS.NUM);
	});
});

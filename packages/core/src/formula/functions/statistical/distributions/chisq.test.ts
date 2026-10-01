import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, stat } from '../test-utils/expect';

describe('CHISQ.DIST / CHISQ.DIST.RT / CHIDIST', () => {
	it('matches Excel', () => {
		expectClose(stat('CHISQ.DIST', [0.5, 1, true]), 0.5204998778130465, 1e-15);
		expectClose(stat('CHISQ.DIST', [2, 3, false]), 0.20755374871029736, 1e-15);
		expectClose(stat('CHISQ.DIST.RT', [18.307, 10]), 0.050000589091398095, 1e-14);
		expectClose(stat('CHIDIST', [18.307, 10]), 0.050000589091398095, 1e-14);
	});

	it.each([
		[100, 10, 1, 5.4497019829205295e-17],
		[1000, 900, 0.9890053910578641, 0.010994608942135848],
		[0.001, 5, 1.6814877189706276e-9, 0.9999999983185123],
		[50, 2, 0.9999999999861121, 1.3887943864964021e-11],
		[300, 250, 0.9834817268421313, 0.01651827315786872],
	])('x = %d, df = %d', (x, df, lower, upper) => {
		expectClose(stat('CHISQ.DIST', [x, df, true]), lower, 1e-13);
		expectClose(stat('CHISQ.DIST.RT', [x, df]), upper, 1e-13);
	});

	it('truncates degrees of freedom and validates', () => {
		expect(stat('CHISQ.DIST.RT', [3, 2.9])).toBe(stat('CHISQ.DIST.RT', [3, 2]));
		expect(stat('CHISQ.DIST', [0, 3, true])).toBe(0);
		expect(stat('CHISQ.DIST', [0, 2, false])).toBe(0.5);
		expect(stat('CHISQ.DIST', [-1, 3, true])).toBe(ERRORS.NUM);
		expect(stat('CHISQ.DIST', [1, 0.5, true])).toBe(ERRORS.NUM);
		expect(stat('CHISQ.DIST.RT', [1, 2e10])).toBe(ERRORS.NUM);
		expect(stat('CHIDIST', [-0.5, 3])).toBe(ERRORS.NUM);
		expect(stat('CHISQ.DIST', [0, 1, false])).toBe(ERRORS.NUM);
	});
});

describe('CHISQ.INV / CHISQ.INV.RT / CHIINV', () => {
	it('matches Excel', () => {
		expectClose(stat('CHISQ.INV', [0.93, 1]), 3.2830202867595357, 1e-14);
		expectClose(stat('CHISQ.INV', [0.6, 2]), 1.8325814637483102, 1e-14);
		expectClose(stat('CHISQ.INV.RT', [0.050001, 10]), 18.306973456961057, 1e-14);
		expectClose(stat('CHISQ.INV.RT', [0.05, 3]), 7.81472790325118, 1e-14);
		expectClose(stat('CHIINV', [0.05, 3]), 7.81472790325118, 1e-14);
	});

	it.each([
		['CHISQ.INV.RT', 0.05, 1, 3.841458820694126],
		['CHISQ.INV.RT', 0.01, 10, 23.20925115895436],
		['CHISQ.INV.RT', 1e-10, 4, 52.66796321106174],
		['CHISQ.INV', 0.001, 200, 143.8427949900008],
		['CHISQ.INV', 0.5, 1, 0.4549364231195728],
		['CHISQ.INV', 0.999, 3, 16.26623619623813],
	])('%s(%d, %d)', (name, p, df, x) => {
		expectClose(stat(name, [p, df]), x, 1e-14);
	});

	it('handles the end points and validates', () => {
		expect(stat('CHISQ.INV', [0, 3])).toBe(0);
		expect(stat('CHISQ.INV.RT', [1, 3])).toBe(0);
		expect(stat('CHISQ.INV', [1, 3])).toBe(ERRORS.NUM);
		expect(stat('CHISQ.INV.RT', [0, 3])).toBe(ERRORS.NUM);
		expect(stat('CHISQ.INV', [1.5, 3])).toBe(ERRORS.NUM);
		expect(stat('CHIINV', [-0.1, 3])).toBe(ERRORS.NUM);
		expect(stat('CHIINV', [0.5, 0])).toBe(ERRORS.NUM);
	});
});

describe('CHISQ.TEST / CHITEST', () => {
	const actual = arr([58, 35], [11, 25], [10, 23]);
	const expected = arr([45.35, 47.65], [17.56, 18.44], [16.09, 16.91]);

	it.each(['CHISQ.TEST', 'CHITEST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [actual, expected]), 0.0003081920170083089, 1e-12);
	});

	it('works on references and uses n - 1 degrees of freedom for a single row', () => {
		const ctx = new MockContext({ A1: 10, B1: 20, C1: 30, A2: 15, B2: 20, C2: 25 });
		// chi^2 = 25/15 + 0 + 25/25 = 2.6667 with df = 2
		expectClose(
			stat('CHISQ.TEST', [ctx.ref('A1:C1'), ctx.ref('A2:C2')], ctx),
			Math.exp(-4 / 3),
			1e-14,
		);
	});

	it('validates shapes and values', () => {
		expect(stat('CHISQ.TEST', [arr([1, 2]), arr([1, 2, 3])])).toBe(ERRORS.NA);
		expect(stat('CHISQ.TEST', [arr([1]), arr([1])])).toBe(ERRORS.NA);
		expect(stat('CHISQ.TEST', [arr([1, 2]), arr([0, 2])])).toBe(ERRORS.DIV0);
		expect(stat('CHISQ.TEST', [arr([1, ERRORS.REF]), arr([1, 2])])).toBe(ERRORS.REF);
	});
});

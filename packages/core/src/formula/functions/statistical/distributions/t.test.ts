import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, stat } from '../test-utils/expect';

describe('T.DIST family', () => {
	it('matches Excel', () => {
		expectClose(stat('T.DIST', [60, 1, true]), 0.9946953263673767, 1e-15);
		expectClose(stat('T.DIST', [8, 3, false]), 0.0007369065209469263, 1e-14);
		expectClose(stat('T.DIST.2T', [1.959999998, 60]), 0.054644929975920874, 1e-14);
		expectClose(stat('T.DIST.RT', [1.959999998, 60]), 0.027322464987960437, 1e-14);
		expectClose(stat('TDIST', [1.959999998, 60, 2]), 0.054644929975920874, 1e-14);
		expectClose(stat('TDIST', [1.959999998, 60, 1]), 0.027322464987960437, 1e-14);
	});

	it.each([
		[-3, 1, 0.10241638234956672],
		[0.5, 2, 0.6666666666666666],
		[-2.5, 10, 0.015723422118304402],
		[4, 4, 0.9919349550499538],
		[10, 30, 0.9999999999771237],
		[1.5, 1000, 0.933034980588957],
		[-40, 5, 9.205981085886477e-8],
	])('T.DIST(%d, %d, TRUE)', (t, df, p) => {
		expectClose(stat('T.DIST', [t, df, true]), p, 1e-14);
		expectClose(stat('T.DIST.RT', [-t, df]), p, 1e-14);
	});

	it('truncates degrees of freedom and validates', () => {
		expect(stat('T.DIST', [1, 2.9, true])).toBe(stat('T.DIST', [1, 2, true]));
		expect(stat('T.DIST', [1, 0.5, true])).toBe(ERRORS.NUM);
		expect(stat('T.DIST.2T', [-1, 5])).toBe(ERRORS.NUM);
		expect(stat('T.DIST.RT', [1, 0])).toBe(ERRORS.NUM);
		expect(stat('TDIST', [-1, 5, 1])).toBe(ERRORS.NUM);
		expect(stat('TDIST', [1, 5, 3])).toBe(ERRORS.NUM);
		expectClose(stat('TDIST', [1, 5, 2.7]), stat('T.DIST.2T', [1, 5]) as number, 1e-15);
		expect(stat('T.DIST.2T', [0, 5])).toBe(1);
	});
});

describe('T.INV family', () => {
	it('matches Excel', () => {
		expectClose(stat('T.INV', [0.75, 2]), 0.816496580927726, 1e-14);
		expectClose(stat('T.INV.2T', [0.546449, 60]), 0.606533075825755, 1e-14);
		expectClose(stat('T.INV.2T', [0.05, 10]), 2.22813885198627, 1e-14);
		expectClose(stat('TINV', [0.05464, 60]), 1.9600411871272458, 1e-14);
	});

	it.each([
		[0.05, 1, 12.706204736174705],
		[0.01, 3, 5.840909309733357],
		[0.5, 25, 0.6844299649042703],
		[1e-8, 5, 71.69350276982087],
	])('T.INV.2T(%d, %d)', (p, df, t) => {
		expectClose(stat('T.INV.2T', [p, df]), t, 1e-14);
		expectClose(stat('TINV', [p, df]), t, 1e-14);
	});

	it('inverts the lower tail on both sides of zero', () => {
		expectClose(stat('T.INV', [0.999, 30]), 3.385184866829305, 1e-14);
		expectClose(stat('T.INV', [0.1, 4]), -1.5332062740589438, 1e-14);
		expect(stat('T.INV', [0.5, 4])).toBe(0);
	});

	it('validates its arguments', () => {
		expect(stat('T.INV', [0, 4])).toBe(ERRORS.NUM);
		expect(stat('T.INV', [1, 4])).toBe(ERRORS.NUM);
		expect(stat('T.INV', [0.5, 0])).toBe(ERRORS.NUM);
		expect(stat('T.INV.2T', [0, 4])).toBe(ERRORS.NUM);
		expect(stat('T.INV.2T', [1.5, 4])).toBe(ERRORS.NUM);
		expect(stat('T.INV.2T', [1, 4])).toBe(0);
		expect(stat('TINV', [0.5, 0.9])).toBe(ERRORS.NUM);
	});
});

describe('CONFIDENCE.T', () => {
	it('matches Excel', () => {
		expectClose(stat('CONFIDENCE.T', [0.05, 1, 50]), 0.2841968554957299, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('CONFIDENCE.T', [0.05, 1, 1])).toBe(ERRORS.DIV0);
		expect(stat('CONFIDENCE.T', [0.05, 1, 0])).toBe(ERRORS.NUM);
		expect(stat('CONFIDENCE.T', [0, 1, 50])).toBe(ERRORS.NUM);
		expect(stat('CONFIDENCE.T', [0.05, 0, 50])).toBe(ERRORS.NUM);
	});
});

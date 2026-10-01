import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, stat } from '../test-utils/expect';

// Expected values: Excel documentation examples, checked against 60-digit references.
describe('NORM.DIST / NORMDIST', () => {
	it.each(['NORM.DIST', 'NORMDIST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [42, 40, 1.5, true]), 0.9087887802741321, 1e-15);
		expectClose(stat(name, [42, 40, 1.5, false]), 0.10934004978399575, 1e-15);
		expectClose(stat(name, [40, 40, 1.5, true]), 0.5, 1e-15);
	});

	it('validates its arguments', () => {
		expect(stat('NORM.DIST', [42, 40, 0, true])).toBe(ERRORS.NUM);
		expect(stat('NORM.DIST', [42, 40, -1, false])).toBe(ERRORS.NUM);
		expect(stat('NORM.DIST', ['x', 40, 1, true])).toBe(ERRORS.VALUE);
		expect(stat('NORM.DIST', [42, 40, 1.5, 'maybe'])).toBe(ERRORS.VALUE);
		expect(stat('NORM.DIST', [42, 40, 1.5])).toBe(ERRORS.VALUE);
	});
});

describe('NORM.S.DIST / NORMSDIST / PHI / GAUSS', () => {
	it('matches Excel', () => {
		expectClose(stat('NORM.S.DIST', [1, true]), 0.841344746068543, 1e-15);
		expectClose(stat('NORM.S.DIST', [1.333333, true]), 0.9087887256040951, 1e-15);
		expectClose(stat('NORM.S.DIST', [1.333333, false]), 0.16401014756936724, 1e-15);
		expectClose(stat('NORMSDIST', [1.333333]), 0.9087887256040951, 1e-15);
		expectClose(stat('NORM.S.DIST', [-20, true]), 2.7536241186062337e-89, 1e-14);
		expectClose(stat('PHI', [0.75]), 0.30113743215480443, 1e-15);
		expectClose(stat('GAUSS', [2]), 0.4772498680518208, 1e-15);
		expectClose(stat('GAUSS', [-0.0001]), -3.989422797365289e-5, 1e-14);
		expect(stat('GAUSS', [0])).toBe(0);
	});
});

describe('normal quantiles', () => {
	it('matches Excel', () => {
		expectClose(stat('NORM.INV', [0.908789, 40, 1.5]), 42.00000200956616, 1e-15);
		expectClose(stat('NORMINV', [0.908789, 40, 1.5]), 42.00000200956616, 1e-15);
		expectClose(stat('NORM.S.INV', [0.908789]), 1.3333346730441074, 1e-15);
		expectClose(stat('NORMSINV', [0.975]), 1.959963984540054, 1e-15);
		expectClose(stat('NORM.S.INV', [1e-300]), -37.0470962993612, 1e-15);
	});

	it.each(['NORM.INV', 'NORMINV'])('%s rejects probabilities outside (0, 1)', (name) => {
		expect(stat(name, [0, 0, 1])).toBe(ERRORS.NUM);
		expect(stat(name, [1, 0, 1])).toBe(ERRORS.NUM);
		expect(stat(name, [0.5, 0, 0])).toBe(ERRORS.NUM);
	});

	it.each(['NORM.S.INV', 'NORMSINV'])('%s rejects probabilities outside (0, 1)', (name) => {
		expect(stat(name, [0])).toBe(ERRORS.NUM);
		expect(stat(name, [1])).toBe(ERRORS.NUM);
		expect(stat(name, [-0.5])).toBe(ERRORS.NUM);
	});
});

describe('lognormal', () => {
	it('matches Excel', () => {
		expectClose(stat('LOGNORM.DIST', [4, 3.5, 1.2, true]), 0.03908355570680047, 1e-14);
		expectClose(stat('LOGNORMDIST', [4, 3.5, 1.2]), 0.03908355570680047, 1e-14);
	});

	it('computes the density and quantile', () => {
		expectClose(stat('LOGNORM.DIST', [4, 3.5, 1.2, false]), 0.017617596681819225, 1e-14);
		expectClose(stat('LOGNORM.INV', [0.039084, 3.5, 1.2]), 4.000025218680635, 1e-14);
		expectClose(stat('LOGINV', [0.039084, 3.5, 1.2]), 4.000025218680635, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('LOGNORM.DIST', [0, 3.5, 1.2, true])).toBe(ERRORS.NUM);
		expect(stat('LOGNORM.DIST', [4, 3.5, 0, true])).toBe(ERRORS.NUM);
		expect(stat('LOGNORMDIST', [-1, 0, 1])).toBe(ERRORS.NUM);
		expect(stat('LOGNORM.INV', [0, 3.5, 1.2])).toBe(ERRORS.NUM);
		expect(stat('LOGINV', [0.5, 3.5, -1])).toBe(ERRORS.NUM);
	});
});

describe('CONFIDENCE / CONFIDENCE.NORM', () => {
	it.each(['CONFIDENCE', 'CONFIDENCE.NORM'])('%s matches Excel', (name) => {
		expectClose(stat(name, [0.05, 2.5, 50]), 0.692951912174839, 1e-15);
		// size is truncated
		expectClose(stat(name, [0.05, 2.5, 50.9]), 0.692951912174839, 1e-15);
	});

	it('validates its arguments', () => {
		expect(stat('CONFIDENCE.NORM', [0, 2.5, 50])).toBe(ERRORS.NUM);
		expect(stat('CONFIDENCE.NORM', [1, 2.5, 50])).toBe(ERRORS.NUM);
		expect(stat('CONFIDENCE.NORM', [0.05, 0, 50])).toBe(ERRORS.NUM);
		expect(stat('CONFIDENCE', [0.05, 2.5, 0.5])).toBe(ERRORS.NUM);
	});
});

describe('FISHER / FISHERINV / STANDARDIZE', () => {
	it('matches Excel', () => {
		expectClose(stat('FISHER', [0.75]), 0.9729550745276566, 1e-15);
		expectClose(stat('FISHERINV', [0.972955]), 0.7499999673941484, 1e-14);
		expectClose(stat('FISHERINV', [stat('FISHER', [-0.3]) as number]), -0.3, 1e-15);
		expectClose(stat('STANDARDIZE', [42, 40, 1.5]), 1.3333333333333333, 1e-15);
	});

	it('validates its arguments', () => {
		expect(stat('FISHER', [1])).toBe(ERRORS.NUM);
		expect(stat('FISHER', [-1])).toBe(ERRORS.NUM);
		expect(stat('STANDARDIZE', [1, 0, 0])).toBe(ERRORS.NUM);
		expect(stat('FISHERINV', ['a'])).toBe(ERRORS.VALUE);
	});
});

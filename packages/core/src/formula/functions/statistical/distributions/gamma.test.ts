import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, stat } from '../test-utils/expect';

describe('GAMMA / GAMMALN', () => {
	it('matches Excel', () => {
		expectClose(stat('GAMMA', [2.5]), 1.329340388179137, 1e-15);
		expectClose(stat('GAMMA', [-3.75]), 0.2678661288614166, 1e-15);
		expect(stat('GAMMA', [5])).toBe(24);
		expectClose(stat('GAMMALN', [4.5]), 2.45373657084244, 1e-14);
		expectClose(stat('GAMMALN', [4]), 1.791759469228055, 1e-15);
		expectClose(stat('GAMMALN.PRECISE', [4]), 1.791759469228055, 1e-15);
		expect(stat('GAMMALN', [1])).toBe(0);
	});

	it('rejects poles and non-positive logarithm arguments', () => {
		expect(stat('GAMMA', [0])).toBe(ERRORS.NUM);
		expect(stat('GAMMA', [-2])).toBe(ERRORS.NUM);
		expect(stat('GAMMA', [172])).toBe(ERRORS.NUM);
		expect(stat('GAMMALN', [0])).toBe(ERRORS.NUM);
		expect(stat('GAMMALN.PRECISE', [-1.5])).toBe(ERRORS.NUM);
		expect(stat('GAMMA', ['x'])).toBe(ERRORS.VALUE);
	});
});

describe('GAMMA.DIST / GAMMADIST / GAMMA.INV / GAMMAINV', () => {
	it.each(['GAMMA.DIST', 'GAMMADIST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [10.00001131, 9, 2, false]), 0.032639130418294, 1e-14);
		expectClose(stat(name, [10.00001131, 9, 2, true]), 0.06809400386978733, 1e-14);
	});

	it('handles other shapes and the origin', () => {
		expectClose(stat('GAMMA.DIST', [5, 2, 3, true]), 0.49633172576650175, 1e-14);
		expectClose(stat('GAMMA.DIST', [0.001, 0.1, 1, false]), 52.6289895128602, 1e-13);
		expectClose(stat('GAMMA.DIST', [150, 100, 1, true]), 0.9999940754596646, 1e-14);
		expect(stat('GAMMA.DIST', [0, 2, 3, true])).toBe(0);
		expect(stat('GAMMA.DIST', [0, 1, 2, false])).toBe(0.5);
		expect(stat('GAMMA.DIST', [0, 0.5, 2, false])).toBe(ERRORS.NUM);
	});

	it.each(['GAMMA.INV', 'GAMMAINV'])('%s matches Excel', (name) => {
		expectClose(stat(name, [0.068094, 9, 2]), 10.000011191437178, 1e-14);
		expectClose(stat(name, [0.5, 3, 1]), 2.6740603137235603, 1e-14);
		expectClose(stat(name, [0.01, 0.5, 2]), 0.00015708785790970197, 1e-13);
		expect(stat(name, [0, 3, 1])).toBe(0);
	});

	it('handles astronomically large shapes quickly', () => {
		const a = 1e20;
		expectClose(
			stat('GAMMA.DIST', [a, a, 1, true]),
			0.5 + 1 / (3 * Math.sqrt(2 * Math.PI * a)),
			1e-15,
		);
		// normal limit: a + z sqrt(a) with z = NORM.S.INV(0.3)
		expectClose(
			stat('GAMMA.INV', [0.3, 1e15, 1]),
			1e15 - 0.5244005127080407 * Math.sqrt(1e15),
			1e-12,
		);
	});

	it('validates its arguments', () => {
		expect(stat('GAMMA.DIST', [-1, 2, 3, true])).toBe(ERRORS.NUM);
		expect(stat('GAMMA.DIST', [1, 0, 3, true])).toBe(ERRORS.NUM);
		expect(stat('GAMMADIST', [1, 2, 0, true])).toBe(ERRORS.NUM);
		expect(stat('GAMMA.INV', [-0.1, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('GAMMA.INV', [1, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('GAMMAINV', [0.5, 2, -3])).toBe(ERRORS.NUM);
	});
});

describe('EXPON.DIST / EXPONDIST / WEIBULL', () => {
	it.each(['EXPON.DIST', 'EXPONDIST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [0.2, 10, true]), 0.8646647167633873, 1e-15);
		expectClose(stat(name, [0.2, 10, false]), 1.353352832366127, 1e-15);
		expect(stat(name, [0, 10, false])).toBe(10);
		expect(stat(name, [-1, 10, true])).toBe(ERRORS.NUM);
		expect(stat(name, [1, 0, true])).toBe(ERRORS.NUM);
	});

	it('keeps relative precision for tiny exponential probabilities', () => {
		expectClose(stat('EXPON.DIST', [1e-20, 1, true]), 1e-20, 1e-15);
	});

	it.each(['WEIBULL', 'WEIBULL.DIST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [105, 20, 100, true]), 0.9295813900692768, 1e-14);
		expectClose(stat(name, [105, 20, 100, false]), 0.03558886402450438, 1e-14);
		expect(stat(name, [0, 1, 2, false])).toBe(0.5);
		expect(stat(name, [0, 0.5, 2, false])).toBe(ERRORS.NUM);
		expect(stat(name, [-1, 2, 2, true])).toBe(ERRORS.NUM);
		expect(stat(name, [1, 0, 2, true])).toBe(ERRORS.NUM);
		expect(stat(name, [1, 2, 0, true])).toBe(ERRORS.NUM);
	});
});

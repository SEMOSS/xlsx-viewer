import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, rowsOf, stat } from '../test-utils/expect';
import { arr } from '../../../test-utils/mock-context';

describe('BINOM.DIST / BINOMDIST', () => {
	it.each(['BINOM.DIST', 'BINOMDIST'])('%s matches Excel exactly for small n', (name) => {
		expect(stat(name, [6, 10, 0.5, false])).toBe(0.205078125);
		expect(stat(name, [3, 10, 0.5, false])).toBe(0.1171875);
		expect(stat(name, [3, 10, 0.5, true])).toBe(0.171875);
	});

	it.each([
		[7, 20, 0.3, 0.16426198521723648, 0.7722717974181605],
		[45, 100, 0.47, 0.07380529452032585, 0.38266956737986146],
		[500, 1000, 0.5, 0.0252250181783608, 0.5126125090891804],
		[2, 2000, 0.001, 0.2708059920477052, 0.6766764387615206],
		[1500, 2000, 0.75, 0.02059757143332068, 0.5085832112651493],
	])('BINOM.DIST(%d, %d, %d)', (k, n, p, pmf, cdf) => {
		expectClose(stat('BINOM.DIST', [k, n, p, false]), pmf, 1e-13);
		expectClose(stat('BINOM.DIST', [k, n, p, true]), cdf, 1e-13);
	});

	it('truncates counts and validates', () => {
		expect(stat('BINOM.DIST', [3.9, 10.2, 0.5, false])).toBe(0.1171875);
		expect(stat('BINOM.DIST', [-1, 10, 0.5, false])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST', [11, 10, 0.5, false])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST', [3, 10, 1.5, false])).toBe(ERRORS.NUM);
		expect(stat('BINOMDIST', [3, 10, -0.1, true])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST', [0, 10, 0, false])).toBe(1);
		expect(stat('BINOM.DIST', [10, 10, 1, true])).toBe(1);
	});

	it('is lifted over an array of successes', () => {
		const out = rowsOf(stat('BINOM.DIST', [arr([0, 1, 2]), 2, 0.5, false]));
		expect(out[0]).toEqual([0.25, 0.5, 0.25]);
	});
});

describe('BINOM.DIST.RANGE', () => {
	it('matches Excel', () => {
		expectClose(stat('BINOM.DIST.RANGE', [60, 0.75, 48]), 0.08397496742904755, 1e-13);
		expectClose(stat('BINOM.DIST.RANGE', [60, 0.75, 45, 50]), 0.5236297934718872, 1e-13);
		expect(stat('BINOM.DIST.RANGE', [10, 0.5, 0, 10])).toBeCloseTo(1, 15);
	});

	it('validates its arguments', () => {
		expect(stat('BINOM.DIST.RANGE', [-1, 0.5, 0])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST.RANGE', [10, 1.2, 3])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST.RANGE', [10, 0.5, 11])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST.RANGE', [10, 0.5, 5, 4])).toBe(ERRORS.NUM);
		expect(stat('BINOM.DIST.RANGE', [10, 0.5, 5, 11])).toBe(ERRORS.NUM);
	});
});

describe('BINOM.INV / CRITBINOM', () => {
	it.each(['BINOM.INV', 'CRITBINOM'])('%s matches Excel', (name) => {
		expect(stat(name, [6, 0.5, 0.75])).toBe(4);
		expect(stat(name, [6, 0.5, 0.65625])).toBe(3);
		expect(stat(name, [100, 0.3, 0.5])).toBe(30);
		expect(stat(name, [10, 0.5, 0])).toBe(0);
		expect(stat(name, [10, 0.5, 1])).toBe(10);
	});

	it('handles very large numbers of trials', () => {
		// np + z sigma - 1/2 rounded up, z = NORM.S.INV(0.975); P(X <= n / 2) = (1 + pmf) / 2
		expect(stat('BINOM.INV', [1e15, 0.5, 0.975])).toBe(500000030989752);
		expectClose(
			stat('BINOM.DIST', [5e14, 1e15, 0.5, true]),
			0.5 + Math.sqrt(2 / (Math.PI * 1e15)) / 2,
			1e-12,
		);
	});

	it('validates its arguments', () => {
		expect(stat('BINOM.INV', [-1, 0.5, 0.5])).toBe(ERRORS.NUM);
		expect(stat('BINOM.INV', [10, 1.5, 0.5])).toBe(ERRORS.NUM);
		expect(stat('CRITBINOM', [10, 0.5, -0.5])).toBe(ERRORS.NUM);
		expect(stat('CRITBINOM', [10, 0.5, 1.5])).toBe(ERRORS.NUM);
	});
});

describe('NEGBINOM.DIST / NEGBINOMDIST', () => {
	it('matches Excel', () => {
		expectClose(stat('NEGBINOM.DIST', [10, 5, 0.25, true]), 0.3135140584781766, 1e-14);
		expectClose(stat('NEGBINOM.DIST', [10, 5, 0.25, false]), 0.05504866037517786, 1e-14);
		expectClose(stat('NEGBINOMDIST', [10, 5, 0.25]), 0.05504866037517786, 1e-14);
		expectClose(stat('NEGBINOM.DIST', [30, 12, 0.6, true]), 0.9999907966264909, 1e-14);
		expectClose(stat('NEGBINOM.DIST', [7, 1, 0.1, false]), 0.04782969, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('NEGBINOM.DIST', [-1, 5, 0.25, true])).toBe(ERRORS.NUM);
		expect(stat('NEGBINOM.DIST', [10, 0, 0.25, true])).toBe(ERRORS.NUM);
		expect(stat('NEGBINOM.DIST', [10, 5, 1.25, true])).toBe(ERRORS.NUM);
		expect(stat('NEGBINOMDIST', [10, 5, -0.25])).toBe(ERRORS.NUM);
	});
});

describe('HYPGEOM.DIST / HYPGEOMDIST', () => {
	it('matches Excel', () => {
		expectClose(stat('HYPGEOM.DIST', [1, 4, 8, 20, true]), 0.46542827657378744, 1e-15);
		expectClose(stat('HYPGEOM.DIST', [1, 4, 8, 20, false]), 0.3632610939112487, 1e-15);
		expectClose(stat('HYPGEOMDIST', [1, 4, 8, 20]), 0.3632610939112487, 1e-15);
		expectClose(stat('HYPGEOM.DIST', [40, 100, 400, 1000, false]), 0.08560998572862107, 1e-13);
		expectClose(stat('HYPGEOM.DIST', [40, 100, 400, 1000, true]), 0.5450660029793728, 1e-13);
		expectClose(stat('HYPGEOM.DIST', [3, 10, 30, 100, true]), 0.6540199886608099, 1e-13);
	});

	it('validates its arguments', () => {
		expect(stat('HYPGEOM.DIST', [5, 4, 8, 20, false])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOM.DIST', [-1, 4, 8, 20, false])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOM.DIST', [0, 15, 8, 20, false])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOM.DIST', [1, 0, 8, 20, false])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOM.DIST', [1, 4, 0, 20, false])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOMDIST', [1, 4, 21, 20])).toBe(ERRORS.NUM);
		expect(stat('HYPGEOMDIST', [1, 25, 8, 20])).toBe(ERRORS.NUM);
	});
});

describe('POISSON / POISSON.DIST', () => {
	it.each(['POISSON', 'POISSON.DIST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [2, 5, true]), 0.12465201948308115, 1e-15);
		expectClose(stat(name, [2, 5, false]), 0.08422433748856833, 1e-15);
		expectClose(stat(name, [100, 100, true]), 0.5265621985299984, 1e-13);
		expectClose(stat(name, [3, 50, true]), 4.269159205144934e-18, 1e-13);
	});

	it('truncates x and validates', () => {
		expectClose(stat('POISSON.DIST', [2.9, 5, false]), 0.08422433748856833, 1e-15);
		expect(stat('POISSON.DIST', [0, 0, false])).toBe(1);
		expect(stat('POISSON.DIST', [-1, 5, true])).toBe(ERRORS.NUM);
		expect(stat('POISSON', [1, -5, true])).toBe(ERRORS.NUM);
	});
});

describe('PERMUT / PERMUTATIONA', () => {
	it('matches Excel', () => {
		expect(stat('PERMUT', [100, 3])).toBe(970200);
		expect(stat('PERMUT', [3, 2])).toBe(6);
		expect(stat('PERMUT', [10.9, 0])).toBe(1);
		expect(stat('PERMUTATIONA', [3, 2])).toBe(9);
		expect(stat('PERMUTATIONA', [2, 2])).toBe(4);
		expect(stat('PERMUTATIONA', [0, 0])).toBe(1);
	});

	it('validates its arguments', () => {
		expect(stat('PERMUT', [0, 0])).toBe(ERRORS.NUM);
		expect(stat('PERMUT', [3, 4])).toBe(ERRORS.NUM);
		expect(stat('PERMUT', [3, -1])).toBe(ERRORS.NUM);
		expect(stat('PERMUT', [1000, 500])).toBe(ERRORS.NUM);
		expect(stat('PERMUTATIONA', [-1, 2])).toBe(ERRORS.NUM);
		expect(stat('PERMUTATIONA', [2, -2])).toBe(ERRORS.NUM);
	});
});

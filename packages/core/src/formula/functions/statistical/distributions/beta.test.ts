import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { expectClose, stat } from '../test-utils/expect';

describe('BETA.DIST / BETADIST', () => {
	it('matches Excel', () => {
		expect(stat('BETA.DIST', [2, 8, 10, true, 1, 3])).toBe(0.6854705810546875);
		expect(stat('BETA.DIST', [2, 8, 10, false, 1, 3])).toBe(1.4837646484375);
		expect(stat('BETADIST', [2, 8, 10, 1, 3])).toBe(0.6854705810546875);
	});

	it.each([
		[0.2, 0.5, 0.5, 0.2951672353008666, 0.7957747154594766],
		[0.9, 2, 5, 0.999945, 0.0026999999999999975],
		[0.01, 0.1, 3, 0.7274341927166095, 7.142534937904182],
		[0.999, 50, 2, 0.9987659096068829, 2.4280023542566886],
		[0.3, 1, 7, 0.9176457, 0.823543],
		[0.75, 3, 1, 0.421875, 1.6875],
		[0.42, 20, 25, 0.3750797876715436, 5.138157576488628],
	])('x = %d on the unit interval, alpha = %d, beta = %d', (x, a, b, cdf, pdf) => {
		expectClose(stat('BETA.DIST', [x, a, b, true]), cdf, 1e-13);
		expectClose(stat('BETA.DIST', [x, a, b, false]), pdf, 1e-13);
		expectClose(stat('BETADIST', [x, a, b]), cdf, 1e-13);
	});

	it('handles the end points', () => {
		expect(stat('BETA.DIST', [0, 2, 3, true])).toBe(0);
		expect(stat('BETA.DIST', [1, 2, 3, true])).toBe(1);
		expect(stat('BETA.DIST', [0, 1, 3, false])).toBe(3);
		expect(stat('BETA.DIST', [0, 0.5, 3, false])).toBe(ERRORS.NUM);
	});

	it('validates its arguments', () => {
		expect(stat('BETA.DIST', [0.5, 0, 3, true])).toBe(ERRORS.NUM);
		expect(stat('BETA.DIST', [0.5, 2, -1, true])).toBe(ERRORS.NUM);
		expect(stat('BETA.DIST', [4, 2, 3, true, 1, 3])).toBe(ERRORS.NUM);
		expect(stat('BETA.DIST', [0.5, 2, 3, true, 1, 1])).toBe(ERRORS.NUM);
		expect(stat('BETADIST', [-0.1, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('BETADIST', ['a', 2, 3])).toBe(ERRORS.VALUE);
	});
});

describe('BETA.INV / BETAINV', () => {
	it.each(['BETA.INV', 'BETAINV'])('%s matches Excel', (name) => {
		expectClose(stat(name, [0.685470581, 8, 10, 1, 3]), 1.9999999999631437, 1e-12);
		expectClose(stat(name, [0.6854705810546875, 8, 10, 1, 3]), 2, 1e-14);
	});

	it.each([
		[0.5, 2, 3, 0.38572756813238956],
		[0.001, 0.5, 4, 2.0897967918172725e-7],
		[0.99, 10, 1.5, 0.9944112827578341],
		[0.3, 30, 40, 0.39698703679444164],
	])('p = %d, alpha = %d, beta = %d', (p, a, b, x) => {
		expectClose(stat('BETA.INV', [p, a, b]), x, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('BETA.INV', [1, 2, 3])).toBe(1);
		expect(stat('BETA.INV', [1, 2, 3, 5, 7])).toBe(7);
		expect(stat('BETA.INV', [0, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('BETA.INV', [1.1, 2, 3])).toBe(ERRORS.NUM);
		expect(stat('BETA.INV', [0.5, 0, 3])).toBe(ERRORS.NUM);
		expect(stat('BETAINV', [0.5, 2, 3, 2, 1])).toBe(ERRORS.NUM);
	});
});

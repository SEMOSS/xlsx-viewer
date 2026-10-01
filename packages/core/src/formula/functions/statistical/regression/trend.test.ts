import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import type { MissingNode } from '../../../ast';
import { arr } from '../../../test-utils/mock-context';
import { column, expectClose, rowsOf, stat } from '../test-utils/expect';

const MISSING: MissingNode = { type: 'missing' };

describe('TREND', () => {
	const months = column(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12);
	const cost = column(
		133890,
		135000,
		135790,
		137300,
		138130,
		139100,
		139900,
		141120,
		141890,
		143230,
		144000,
		145290,
	);

	it('matches the Excel example for fitted values', () => {
		const out = rowsOf(stat('TREND', [cost, months]));
		const expected = [
			133953.33333333334, 134971.51515151514, 135989.69696969696, 137007.87878787878,
			138026.0606060606, 139044.24242424243, 140062.42424242425, 141080.60606060605,
			142098.78787878787, 143116.9696969697, 144135.15151515152, 145153.33333333334,
		];
		expect(out).toHaveLength(12);
		expected.forEach((v, i) => expectClose(out[i][0], v, 1e-14));
	});

	it('predicts new values in the shape of new_x', () => {
		const out = rowsOf(stat('TREND', [cost, months, column(13, 14, 15, 16, 17)]));
		const expected = [
			146171.51515151514, 147189.69696969696, 148207.87878787878, 149226.0606060606,
			150244.24242424243,
		];
		expected.forEach((v, i) => expectClose(out[i][0], v, 1e-14));
		const wide = rowsOf(stat('TREND', [cost, months, arr([13, 14])]));
		expect(wide).toHaveLength(1);
		expectClose(wide[0][1], 147189.69696969696, 1e-14);
	});

	it('handles omitted known_x, no constant and several variables', () => {
		expect(rowsOf(stat('TREND', [column(3, 5, 7), MISSING, column(4, 5)]))).toEqual([[9], [11]]);
		expect(rowsOf(stat('TREND', [column(2, 4, 6), column(1, 2, 3), column(10), false]))).toEqual([
			[20],
		]);
		const y = column(3, 5, 6, 9, 11, 14);
		const x = arr([1, 2], [2, 1], [3, 4], [4, 3], [5, 7], [6, 5]);
		const out = rowsOf(stat('TREND', [y, x, arr([7, 6])]));
		expectClose(
			out[0][0],
			0.4934354485776805 + 2.3544857768052516 * 7 - 0.20021881838074398 * 6,
			1e-13,
		);
	});

	it('validates new_x against the variables', () => {
		const y = column(3, 5, 6, 9);
		const x = arr([1, 2], [2, 1], [3, 4], [4, 3]);
		expect(stat('TREND', [y, x, arr([1, 2, 3])])).toBe(ERRORS.REF);
		expect(stat('TREND', [y, x, arr(['a', 2])])).toBe(ERRORS.VALUE);
		expect(stat('TREND', [column(1, 2), column(1, 2, 3)])).toBe(ERRORS.REF);
	});
});

describe('GROWTH', () => {
	const units = column(33100, 47300, 69000, 102000, 150000, 220000);
	const months = column(11, 12, 13, 14, 15, 16);

	it('matches the Excel example', () => {
		const fitted = rowsOf(stat('GROWTH', [units, months]));
		const expected = [
			32618.2037735398, 47729.42261474788, 69841.30085621761, 102197.07337883263,
			149542.48674004624, 218821.8762145961,
		];
		expected.forEach((v, i) => expectClose(fitted[i][0], v, 1e-13));
		const next = rowsOf(stat('GROWTH', [units, months, column(17, 18)]));
		expectClose(next[0][0], 320196.7183634732, 1e-13);
		expectClose(next[1][0], 468536.0541840495, 1e-13);
	});

	it('rejects non-positive values', () => {
		expect(stat('GROWTH', [column(1, 0, 3), column(1, 2, 3)])).toBe(ERRORS.NUM);
		expect(stat('GROWTH', [column(1, 2, 3), column(1, 2, 3), column(4), 'x'])).toBe(ERRORS.VALUE);
	});
});

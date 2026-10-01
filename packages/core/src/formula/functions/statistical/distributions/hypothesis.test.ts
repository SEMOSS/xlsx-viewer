import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, stat } from '../test-utils/expect';

const A = arr([3, 4, 5, 8, 9, 1, 2, 4, 5]);
const B = arr([6, 19, 3, 2, 14, 4, 5, 17, 1]);

describe('T.TEST / TTEST', () => {
	it.each(['T.TEST', 'TTEST'])('%s matches Excel for the paired test', (name) => {
		expectClose(stat(name, [A, B, 2, 1]), 0.1960157849252821, 1e-13);
		expectClose(stat(name, [A, B, 1, 1]), 0.1960157849252821 / 2, 1e-13);
	});

	it('runs the equal- and unequal-variance tests (Welch df not truncated)', () => {
		expectClose(stat('T.TEST', [A, B, 2, 2]), 0.19199588676039622, 1e-13);
		expectClose(stat('T.TEST', [A, B, 1, 2]), 0.09599794338019811, 1e-13);
		expectClose(stat('T.TEST', [A, B, 2, 3]), 0.20229392336867788, 1e-13);
	});

	it('pairs observations positionally and skips non-numeric pairs', () => {
		const ctx = new MockContext({
			A1: 3,
			A2: 'x',
			A3: 5,
			A4: 8,
			B1: 6,
			B2: 19,
			B3: 3,
			B4: true,
		});
		// Only (3, 6) and (5, 3) are used.
		const t = Math.abs(-0.5) / (Math.sqrt(12.5) / Math.sqrt(2));
		const expected = stat('T.DIST.2T', [t, 1]) as number;
		expectClose(stat('T.TEST', [ctx.ref('A1:A4'), ctx.ref('B1:B4'), 2, 1], ctx), expected, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('T.TEST', [A, B, 3, 1])).toBe(ERRORS.NUM);
		expect(stat('T.TEST', [A, B, 2, 4])).toBe(ERRORS.NUM);
		expect(stat('T.TEST', [A, B, 0, 1])).toBe(ERRORS.NUM);
		expect(stat('T.TEST', [A, arr([1, 2, 3]), 2, 1])).toBe(ERRORS.NA);
		expect(stat('T.TEST', [arr([1]), arr([2]), 2, 2])).toBe(ERRORS.DIV0);
		expect(stat('T.TEST', [arr([1, 1]), arr([2, 2]), 2, 2])).toBe(ERRORS.DIV0);
		expect(stat('T.TEST', [arr([1, ERRORS.NA]), B, 2, 2])).toBe(ERRORS.NA);
		expect(stat('TTEST', [A, B, 'x', 1])).toBe(ERRORS.VALUE);
	});
});

describe('F.TEST / FTEST', () => {
	it.each(['F.TEST', 'FTEST'])('%s matches Excel', (name) => {
		expectClose(
			stat(name, [arr([6, 7, 9, 15, 21]), arr([20, 28, 31, 38, 40])]),
			0.6483178467861745,
			1e-13,
		);
	});

	it('validates its arguments', () => {
		expect(stat('F.TEST', [arr([1]), arr([1, 2])])).toBe(ERRORS.DIV0);
		expect(stat('F.TEST', [arr([3, 3, 3]), arr([1, 2])])).toBe(ERRORS.DIV0);
		// Text is skipped: variances 2 and 0.5, F = 4 on (1, 1) df, P(F > 4) = 1 - (2 / pi) atan(2).
		expectClose(
			stat('F.TEST', [arr([1, 'a', 3]), arr([1, 2])]),
			2 * (1 - (2 / Math.PI) * Math.atan(2)),
			1e-14,
		);
	});
});

describe('Z.TEST / ZTEST', () => {
	const data = arr([3, 6, 7, 8, 6, 5, 4, 2, 1, 9]);

	it.each(['Z.TEST', 'ZTEST'])('%s matches Excel', (name) => {
		expectClose(stat(name, [data, 4]), 0.09057419685136377, 1e-14);
		expectClose(stat(name, [data, 6]), 0.8630433891295299, 1e-14);
		expectClose(stat(name, [data, 4, 2]), 0.04099516050019149, 1e-14);
	});

	it('validates its arguments', () => {
		expect(stat('Z.TEST', [arr(['a']), 4])).toBe(ERRORS.NA);
		expect(stat('Z.TEST', [arr([5]), 4])).toBe(ERRORS.DIV0);
		expect(stat('Z.TEST', [arr([5, 5]), 4])).toBe(ERRORS.DIV0);
		expect(stat('Z.TEST', [data, 4, 0])).toBe(ERRORS.NUM);
		expect(stat('Z.TEST', [data, 'x'])).toBe(ERRORS.VALUE);
	});
});

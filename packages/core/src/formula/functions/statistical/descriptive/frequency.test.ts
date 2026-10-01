import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, rowsOf, stat } from '../test-utils/expect';

describe('FREQUENCY', () => {
	const scores = arr([79, 85, 78, 85, 50, 81, 95, 88, 97]);

	it('matches Excel with one more count than bins', () => {
		expect(rowsOf(stat('FREQUENCY', [scores, arr([70, 79, 89])]))).toEqual([[1], [2], [4], [2]]);
	});

	it('keeps the original bin order and gives duplicate bins nothing', () => {
		expect(rowsOf(stat('FREQUENCY', [scores, arr([89, 70, 79])]))).toEqual([[4], [1], [2], [2]]);
		expect(rowsOf(stat('FREQUENCY', [scores, arr([79, 79])]))).toEqual([[3], [0], [6]]);
	});

	it('handles empty bins and empty data', () => {
		expect(rowsOf(stat('FREQUENCY', [scores, arr(['none'])]))).toEqual([[9]]);
		expect(rowsOf(stat('FREQUENCY', [arr(['x']), arr([1, 2])]))).toEqual([[0], [0], [0]]);
	});

	it('reads references, ignoring text and blanks, and propagates errors', () => {
		const ctx = new MockContext({ A1: 1, A2: 'x', A3: 5, A5: 9, B1: 4 });
		expect(rowsOf(stat('FREQUENCY', [ctx.ref('A1:A5'), ctx.ref('B1')], ctx))).toEqual([[1], [2]]);
		expect(stat('FREQUENCY', [arr([1, ERRORS.REF]), arr([1])])).toBe(ERRORS.REF);
	});
});

describe('PROB', () => {
	const xs = arr([0, 1, 2, 3]);
	const ps = arr([0.2, 0.3, 0.1, 0.4]);

	it('matches Excel', () => {
		expect(stat('PROB', [xs, ps, 2])).toBe(0.1);
		expectClose(stat('PROB', [xs, ps, 1, 3]), 0.8, 1e-15);
		expect(stat('PROB', [xs, ps, 5])).toBe(0);
	});

	it('validates the probabilities and shapes', () => {
		expect(stat('PROB', [xs, arr([0.2, 0.3, 0.1, 0.3]), 1])).toBe(ERRORS.NUM);
		expect(stat('PROB', [xs, arr([0.2, 0.3, 0.6, -0.1]), 1])).toBe(ERRORS.NUM);
		expect(stat('PROB', [xs, arr([0.5, 0.5, 0, 0]), 1])).toBe(ERRORS.NUM);
		expect(stat('PROB', [xs, arr([0.5, 0.5]), 1])).toBe(ERRORS.NA);
		expect(stat('PROB', [xs, ps, 'a'])).toBe(ERRORS.VALUE);
	});
});

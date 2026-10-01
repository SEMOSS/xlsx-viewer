import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, stat } from '../test-utils/expect';

// A1:A6 = 10, 7, "text", TRUE, (blank), 9 ; B1 = #N/A
const ctx = () =>
	new MockContext({ A1: 10, A2: 7, A3: 'text', A4: true, A6: 9, B1: ERRORS.NA, C1: '' });

describe('AVERAGE / AVERAGEA', () => {
	it('ignores text, logical values and blanks inside references', () => {
		const c = ctx();
		expectClose(stat('AVERAGE', [c.ref('A1:A6')], c), 26 / 3, 1e-15);
		expect(stat('AVERAGE', [c.ref('A1:A6'), 2], c)).toBe(7);
	});

	it('coerces typed arguments', () => {
		expect(stat('AVERAGE', [1, '3', true])).toBeCloseTo(5 / 3, 15);
		expect(stat('AVERAGE', [1, 'abc'])).toBe(ERRORS.VALUE);
		expect(stat('AVERAGE', [arr([1, true, '3'])])).toBe(1);
		expect(stat('AVERAGE', [10, 7, 9, 27, 2])).toBe(11);
	});

	it('propagates errors and rejects empty data', () => {
		const c = ctx();
		expect(stat('AVERAGE', [c.ref('A1:B1')], c)).toBe(ERRORS.NA);
		expect(stat('AVERAGE', [c.ref('D1:D5')], c)).toBe(ERRORS.DIV0);
		expect(stat('AVERAGE', ['x'])).toBe(ERRORS.VALUE);
	});

	it('counts TRUE as 1 and text as 0 in AVERAGEA', () => {
		const c = ctx();
		// 10 + 7 + 0 + 1 + 9 over 5 values (the blank is skipped)
		expect(stat('AVERAGEA', [c.ref('A1:A6')], c)).toBe(5.4);
		expect(stat('AVERAGEA', [c.ref('C1')], c)).toBe(0);
		expect(stat('AVERAGEA', [c.ref('D1:D3')], c)).toBe(ERRORS.DIV0);
	});
});

describe('MEDIAN', () => {
	it('matches Excel', () => {
		expect(stat('MEDIAN', [arr([1, 2, 3, 4, 5])])).toBe(3);
		expect(stat('MEDIAN', [arr([1, 2, 3, 4, 5, 6])])).toBe(3.5);
		expect(stat('MEDIAN', [5, 1, 3])).toBe(3);
		const c = ctx();
		expect(stat('MEDIAN', [c.ref('A1:A6')], c)).toBe(9);
		expect(stat('MEDIAN', [c.ref('D1:D2')], c)).toBe(ERRORS.NUM);
	});
});

describe('GEOMEAN / HARMEAN', () => {
	it('matches Excel', () => {
		expectClose(stat('GEOMEAN', [arr([4, 5, 8, 7, 11, 4, 3])]), 5.476986969656962, 1e-15);
		expectClose(stat('HARMEAN', [arr([4, 5, 8, 7, 11, 4, 3])]), 5.028375962061728, 1e-15);
		expect(stat('GEOMEAN', [2, 8])).toBe(4);
		expect(stat('HARMEAN', [1, 1])).toBe(1);
	});

	it('stays exact far outside the floating-point exponent range of a product', () => {
		expectClose(stat('GEOMEAN', [1e300, 1e300, 1e-300]), 1e100, 1e-15);
		expectClose(stat('GEOMEAN', [arr(new Array(400).fill(1e300))]), 1e300, 1e-14);
	});

	it('rejects non-positive values and empty data', () => {
		expect(stat('GEOMEAN', [1, 0])).toBe(ERRORS.NUM);
		expect(stat('GEOMEAN', [arr(['a'])])).toBe(ERRORS.NUM);
		expect(stat('HARMEAN', [1, -1])).toBe(ERRORS.NUM);
		expect(stat('HARMEAN', [arr(['a'])])).toBe(ERRORS.NA);
	});
});

describe('AVEDEV / DEVSQ / TRIMMEAN', () => {
	it('matches Excel', () => {
		expectClose(stat('AVEDEV', [arr([4, 5, 6, 7, 5, 4, 3])]), 1.0204081632653061, 1e-15);
		expect(stat('DEVSQ', [arr([4, 5, 8, 7, 11, 4, 3])])).toBe(48);
		expectClose(
			stat('TRIMMEAN', [arr([4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3]), 0.2]),
			3.7777777777777777,
			1e-15,
		);
	});

	it('is accurate for data with a large offset (two-pass algorithm)', () => {
		const big = [1e9 + 4, 1e9 + 7, 1e9 + 13, 1e9 + 16];
		expect(stat('DEVSQ', [arr(big)])).toBe(90);
		expect(stat('AVEDEV', [arr(big)])).toBe(4.5);
	});

	it('trims floor(n * percent / 2) points from each end', () => {
		expect(stat('TRIMMEAN', [arr([1, 2, 3, 4, 100]), 0.4])).toBe(3);
		expect(stat('TRIMMEAN', [arr([1, 2, 3, 4, 100]), 0.39])).toBe(22);
		expect(stat('TRIMMEAN', [arr([1, 2, 3]), 0])).toBe(2);
	});

	it('validates its arguments', () => {
		expect(stat('AVEDEV', [arr(['a'])])).toBe(ERRORS.NUM);
		expect(stat('DEVSQ', [arr(['a'])])).toBe(ERRORS.NUM);
		expect(stat('TRIMMEAN', [arr([1, 2]), 1])).toBe(ERRORS.NUM);
		expect(stat('TRIMMEAN', [arr([1, 2]), -0.1])).toBe(ERRORS.NUM);
		expect(stat('TRIMMEAN', [arr(['a']), 0.1])).toBe(ERRORS.NUM);
		expect(stat('TRIMMEAN', [arr([1, ERRORS.DIV0]), 0.1])).toBe(ERRORS.DIV0);
	});
});

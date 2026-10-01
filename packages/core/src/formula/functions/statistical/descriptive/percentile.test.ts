import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { rowsOf, stat } from '../test-utils/expect';

const EXC = arr([1, 2, 3, 6, 6, 6, 7, 8, 9]);
const RANKED = arr([13, 12, 11, 8, 4, 3, 2, 1, 1, 1]);

describe('PERCENTILE family', () => {
	it.each(['PERCENTILE', 'PERCENTILE.INC'])('%s matches Excel', (name) => {
		expect(stat(name, [arr([1, 3, 2, 4]), 0.3])).toBe(1.9);
		expect(stat(name, [arr([1, 3, 2, 4]), 0])).toBe(1);
		expect(stat(name, [arr([1, 3, 2, 4]), 1])).toBe(4);
		expect(stat(name, [arr([1, 3, 2, 4]), 1.1])).toBe(ERRORS.NUM);
		expect(stat(name, [arr([1, 3, 2, 4]), -0.1])).toBe(ERRORS.NUM);
	});

	it('matches Excel for PERCENTILE.EXC', () => {
		expect(stat('PERCENTILE.EXC', [EXC, 0.25])).toBe(2.5);
		expect(stat('PERCENTILE.EXC', [EXC, 0.1])).toBe(1);
		expect(stat('PERCENTILE.EXC', [EXC, 0.9])).toBe(9);
		expect(stat('PERCENTILE.EXC', [EXC, 0])).toBe(ERRORS.NUM);
		expect(stat('PERCENTILE.EXC', [EXC, 0.01])).toBe(ERRORS.NUM);
		expect(stat('PERCENTILE.EXC', [EXC, 0.95])).toBe(ERRORS.NUM);
		expect(stat('PERCENTILE.EXC', [EXC, 2])).toBe(ERRORS.NUM);
	});

	it('rejects empty data and lifts over k', () => {
		expect(stat('PERCENTILE.INC', [arr(['a']), 0.5])).toBe(ERRORS.NUM);
		expect(rowsOf(stat('PERCENTILE.INC', [arr([1, 2, 3, 4, 5]), arr([0, 0.5, 1])]))).toEqual([
			[1, 3, 5],
		]);
	});
});

describe('QUARTILE family', () => {
	const data = arr([1, 2, 4, 7, 8, 9, 10, 12]);

	it.each(['QUARTILE', 'QUARTILE.INC'])('%s matches Excel', (name) => {
		expect(stat(name, [data, 1])).toBe(3.5);
		expect(stat(name, [data, 0])).toBe(1);
		expect(stat(name, [data, 2])).toBe(7.5);
		expect(stat(name, [data, 4])).toBe(12);
		expect(stat(name, [data, 3.9])).toBe(9.25);
		expect(stat(name, [data, 5])).toBe(ERRORS.NUM);
		expect(stat(name, [data, -1])).toBe(ERRORS.NUM);
	});

	it('matches Excel for QUARTILE.EXC', () => {
		const exc = arr([6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49]);
		expect(stat('QUARTILE.EXC', [exc, 1])).toBe(15);
		expect(stat('QUARTILE.EXC', [exc, 3])).toBe(43);
		expect(stat('QUARTILE.EXC', [exc, 2])).toBe(40);
		expect(stat('QUARTILE.EXC', [exc, 0])).toBe(ERRORS.NUM);
		expect(stat('QUARTILE.EXC', [exc, 4])).toBe(ERRORS.NUM);
	});
});

describe('PERCENTRANK family', () => {
	it.each(['PERCENTRANK', 'PERCENTRANK.INC'])('%s matches Excel (truncating)', (name) => {
		expect(stat(name, [RANKED, 2])).toBe(0.333);
		expect(stat(name, [RANKED, 4])).toBe(0.555);
		expect(stat(name, [RANKED, 8])).toBe(0.666);
		expect(stat(name, [RANKED, 5])).toBe(0.583);
		expect(stat(name, [RANKED, 1])).toBe(0);
		expect(stat(name, [RANKED, 13])).toBe(1);
	});

	it('interpolates from the last occurrence of the lower value', () => {
		// 1.5 lies between the third 1 (position 2 of 9) and 2 (position 3 of 9)
		expect(stat('PERCENTRANK.INC', [RANKED, 1.5, 4])).toBe(0.2777);
	});

	it('matches Excel for PERCENTRANK.EXC', () => {
		expect(stat('PERCENTRANK.EXC', [EXC, 7])).toBe(0.7);
		expect(stat('PERCENTRANK.EXC', [EXC, 5.43])).toBe(0.381);
		expect(stat('PERCENTRANK.EXC', [EXC, 5.43, 1])).toBe(0.3);
		expect(stat('PERCENTRANK.EXC', [EXC, 1])).toBe(0.1);
	});

	it('validates its arguments', () => {
		expect(stat('PERCENTRANK.INC', [RANKED, 0.5])).toBe(ERRORS.NA);
		expect(stat('PERCENTRANK.INC', [RANKED, 14])).toBe(ERRORS.NA);
		expect(stat('PERCENTRANK.INC', [RANKED, 2, 0])).toBe(ERRORS.NUM);
		expect(stat('PERCENTRANK.EXC', [arr(['a']), 2])).toBe(ERRORS.NUM);
		expect(stat('PERCENTRANK', [arr([5]), 5])).toBe(1);
	});

	it('reads references', () => {
		const ctx = new MockContext({ A1: 10, A2: 'skip', A3: 20, A4: 30 });
		expect(stat('PERCENTRANK.INC', [ctx.ref('A1:A4'), 20], ctx)).toBe(0.5);
	});
});

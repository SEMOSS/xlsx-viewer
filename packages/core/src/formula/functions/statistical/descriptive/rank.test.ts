import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { rowsOf, stat } from '../test-utils/expect';

describe('RANK / RANK.EQ / RANK.AVG', () => {
	const ctx = () => new MockContext({ A1: 7, A2: 3.5, A3: 3.5, A4: 1, A5: 2, A6: 'x' });

	it.each(['RANK', 'RANK.EQ'])('%s matches Excel', (name) => {
		const c = ctx();
		expect(stat(name, [7, c.ref('A1:A5'), 1], c)).toBe(5);
		expect(stat(name, [2, c.ref('A1:A5')], c)).toBe(4);
		expect(stat(name, [3.5, c.ref('A1:A5'), 1], c)).toBe(3);
		expect(stat(name, [3.5, c.ref('A1:A6')], c)).toBe(2);
	});

	it('averages tied ranks in RANK.AVG', () => {
		expect(stat('RANK.AVG', [94, arr([89, 88, 92, 101, 94, 97, 95])])).toBe(4);
		const c = ctx();
		expect(stat('RANK.AVG', [3.5, c.ref('A1:A5')], c)).toBe(2.5);
		expect(stat('RANK.AVG', [3.5, c.ref('A1:A5'), 1], c)).toBe(3.5);
	});

	it('is #N/A when the number is not in the list', () => {
		const c = ctx();
		expect(stat('RANK.EQ', [4, c.ref('A1:A5')], c)).toBe(ERRORS.NA);
		expect(stat('RANK.AVG', [4, c.ref('A1:A5')], c)).toBe(ERRORS.NA);
	});

	it('is lifted over the number and propagates errors', () => {
		const c = ctx();
		expect(rowsOf(stat('RANK.EQ', [arr([7, 1]), c.ref('A1:A5')], c))).toEqual([[1, 5]]);
		expect(stat('RANK.EQ', [1, arr([1, ERRORS.NA])])).toBe(ERRORS.NA);
		expect(stat('RANK', ['x', arr([1])])).toBe(ERRORS.VALUE);
	});
});

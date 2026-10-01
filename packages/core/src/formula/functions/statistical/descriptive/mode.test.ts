import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { rowsOf, stat } from '../test-utils/expect';
import { modes } from './mode';

describe('MODE / MODE.SNGL', () => {
	it.each(['MODE', 'MODE.SNGL'])('%s matches Excel', (name) => {
		expect(stat(name, [arr([5.6, 4, 4, 3, 2, 4])])).toBe(4);
		expect(stat(name, [1, 2, 2, 3])).toBe(2);
	});

	it('returns the first mode in data order on ties', () => {
		expect(stat('MODE', [arr([2, 1, 1, 2])])).toBe(2);
		expect(stat('MODE.SNGL', [arr([1, 3, 3, 1, 2])])).toBe(1);
	});

	it('is #N/A without a repeated value', () => {
		expect(stat('MODE', [arr([1, 2, 3])])).toBe(ERRORS.NA);
		expect(stat('MODE', [arr(['a', 'a'])])).toBe(ERRORS.NA);
		expect(stat('MODE.SNGL', [arr([1, ERRORS.DIV0])])).toBe(ERRORS.DIV0);
	});
});

describe('MODE.MULT', () => {
	it('matches Excel with a vertical result', () => {
		const out = rowsOf(stat('MODE.MULT', [arr([1, 2, 3, 4, 3, 2, 1, 2, 3, 5, 6, 1])]));
		expect(out).toEqual([[1], [2], [3]]);
	});

	it('reads references and returns #N/A without repeats', () => {
		const ctx = new MockContext({ A1: 7, A2: 'x', A3: 7, A4: 8 });
		expect(rowsOf(stat('MODE.MULT', [ctx.ref('A1:A4')], ctx))).toEqual([[7]]);
		expect(stat('MODE.MULT', [arr([1, 2])])).toBe(ERRORS.NA);
	});

	it('keeps first-occurrence order', () => {
		expect(modes([9, 1, 9, 1, 5])).toEqual([9, 1]);
		expect(modes([0, -0, 1])).toEqual([0]);
		expect(modes([])).toEqual([]);
	});
});

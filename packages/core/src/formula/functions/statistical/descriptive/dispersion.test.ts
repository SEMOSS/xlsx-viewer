import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { expectClose, stat } from '../test-utils/expect';

// Excel's STDEV documentation sample (breaking strength of tools).
const TOOLS = arr([1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299]);

describe('variance and standard deviation', () => {
	it.each([
		['VAR', 754.2666666666667],
		['VAR.S', 754.2666666666667],
		['VARP', 678.84],
		['VAR.P', 678.84],
		['STDEV', 27.463915719843495],
		['STDEV.S', 27.463915719843495],
		['STDEVP', 26.054558142482477],
		['STDEV.P', 26.054558142482477],
	])('%s matches Excel', (name, expected) => {
		expectClose(stat(name, [TOOLS]), expected, 1e-15);
	});

	it('needs two points for the sample statistics and one for the population ones', () => {
		expect(stat('VAR.S', [5])).toBe(ERRORS.DIV0);
		expect(stat('STDEV', [arr(['a', 5])])).toBe(ERRORS.DIV0);
		expect(stat('VAR.P', [5])).toBe(0);
		expect(stat('STDEV.P', [arr(['a'])])).toBe(ERRORS.DIV0);
	});

	it('counts logical values and text inside references for the A forms', () => {
		const ctx = new MockContext({ A1: 2, A2: true, A3: 'x', A4: 4, A5: false });
		// values 2, 1, 0, 4, 0: mean 1.4, sum of squares of deviations 11.2
		expectClose(stat('VARA', [ctx.ref('A1:A5')], ctx), 2.8, 1e-15);
		expectClose(stat('VARPA', [ctx.ref('A1:A5')], ctx), 2.24, 1e-15);
		expectClose(stat('STDEVA', [ctx.ref('A1:A5')], ctx), Math.sqrt(2.8), 1e-15);
		expectClose(stat('STDEVPA', [ctx.ref('A1:A5')], ctx), Math.sqrt(2.24), 1e-15);
		// the plain forms see only 2 and 4
		expect(stat('VAR', [ctx.ref('A1:A5')], ctx)).toBe(2);
	});

	it('is exact for data with a large offset', () => {
		expect(stat('VAR.S', [arr([1e9 + 4, 1e9 + 7, 1e9 + 13, 1e9 + 16])])).toBe(30);
		expect(stat('VAR.P', [arr([0.1, 0.1, 0.1])])).toBe(0);
	});

	it('propagates errors', () => {
		expect(stat('STDEV.S', [arr([1, ERRORS.NUM, 3])])).toBe(ERRORS.NUM);
		expect(stat('VAR', [1, 'abc'])).toBe(ERRORS.VALUE);
	});
});

describe('shape statistics', () => {
	const data = arr([3, 4, 5, 2, 3, 4, 5, 6, 4, 7]);

	it('matches Excel', () => {
		expectClose(stat('KURT', [data]), -0.15179963720841627, 1e-14);
		expectClose(stat('SKEW', [data]), 0.3595430714067974, 1e-14);
		expectClose(stat('SKEW.P', [data]), 0.3031933393541438, 1e-14);
	});

	it('is zero for symmetric data', () => {
		expect(stat('SKEW', [arr([1, 2, 3])])).toBe(0);
		expect(stat('SKEW.P', [arr([1, 2, 3, 4])])).toBe(0);
	});

	it('needs enough points and a non-zero spread', () => {
		expect(stat('KURT', [arr([1, 2, 3])])).toBe(ERRORS.DIV0);
		expect(stat('KURT', [arr([2, 2, 2, 2])])).toBe(ERRORS.DIV0);
		expect(stat('SKEW', [arr([1, 2])])).toBe(ERRORS.DIV0);
		expect(stat('SKEW', [arr([5, 5, 5])])).toBe(ERRORS.DIV0);
		expect(stat('SKEW.P', [arr([1, 2])])).toBe(ERRORS.DIV0);
	});
});

describe('STANDARDIZE', () => {
	it('matches Excel', () => {
		expectClose(stat('STANDARDIZE', [42, 40, 1.5]), 1.3333333333333333, 1e-15);
		expect(stat('STANDARDIZE', [1, 2, -1])).toBe(ERRORS.NUM);
	});
});

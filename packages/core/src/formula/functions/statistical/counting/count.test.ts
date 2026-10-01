import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import type { MissingNode } from '../../../ast';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { stat } from '../test-utils/expect';

// Excel's COUNT sample: A2 "Data", A3 a date serial, A4 blank, A5 19, A6 22.24, A7 TRUE, A8 #DIV/0!
const sample = () =>
	new MockContext({ A2: 'Data', A3: 39790, A5: 19, A6: 22.24, A7: true, A8: ERRORS.DIV0, B1: '' });

describe('COUNT', () => {
	it('counts numbers in references only', () => {
		const ctx = sample();
		expect(stat('COUNT', [ctx.ref('A2:A8')], ctx)).toBe(3);
		expect(stat('COUNT', [ctx.ref('A5:A8')], ctx)).toBe(2);
		expect(stat('COUNT', [ctx.ref('A2:A8'), 2], ctx)).toBe(4);
		expect(stat('COUNT', [ctx.ref('A:A')], ctx)).toBe(3);
	});

	it('counts typed numbers, logical values and numeric text', () => {
		expect(stat('COUNT', [1, '2', true, 'abc', ERRORS.NA])).toBe(3);
		expect(stat('COUNT', [arr([1, '2', true, null])])).toBe(1);
		const missing: MissingNode = { type: 'missing' };
		expect(stat('COUNT', [1, missing])).toBe(1);
	});
});

describe('COUNTA', () => {
	it('counts every non-empty value, including errors and empty text', () => {
		const ctx = sample();
		expect(stat('COUNTA', [ctx.ref('A2:A8')], ctx)).toBe(6);
		expect(stat('COUNTA', [ctx.ref('B1')], ctx)).toBe(1);
		expect(stat('COUNTA', [ctx.ref('C1:C10')], ctx)).toBe(0);
		expect(stat('COUNTA', [1, 'x', ERRORS.NA, ''])).toBe(4);
		expect(stat('COUNTA', [arr([1, null, 'a'])])).toBe(2);
	});
});

describe('COUNTBLANK', () => {
	it('counts empty cells and empty text', () => {
		const ctx = sample();
		// A2:A8 has one empty cell (A4)
		expect(stat('COUNTBLANK', [ctx.ref('A2:A8')], ctx)).toBe(1);
		expect(stat('COUNTBLANK', [ctx.ref('B1:B3')], ctx)).toBe(3);
		expect(stat('COUNTBLANK', [ctx.ref('C1:D5')], ctx)).toBe(10);
		expect(stat('COUNTBLANK', [ctx.ref('A:A')], ctx)).toBe(1048576 - 6);
		expect(stat('COUNTBLANK', [arr([1, '', null])])).toBe(2);
	});

	it('requires a range', () => {
		expect(stat('COUNTBLANK', [5])).toBe(ERRORS.VALUE);
		expect(stat('COUNTBLANK', [ERRORS.REF])).toBe(ERRORS.REF);
	});
});

import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);
const MAX = 2 ** 48 - 1;

describe('BITAND / BITOR / BITXOR', () => {
	it('match the documented examples', () => {
		expect(f('BITAND', 1, 5)).toBe(1);
		expect(f('BITAND', 13, 25)).toBe(9);
		expect(f('BITOR', 23, 10)).toBe(31);
		expect(f('BITXOR', 5, 3)).toBe(6);
	});

	it('work on the full 48-bit range', () => {
		expect(f('BITAND', MAX, 2 ** 47)).toBe(2 ** 47);
		expect(f('BITOR', 2 ** 47, 1)).toBe(2 ** 47 + 1);
		expect(f('BITXOR', MAX, MAX)).toBe(0);
		expect(f('BITXOR', MAX, 0)).toBe(MAX);
		expect(f('BITAND', '12', 10)).toBe(8);
	});

	it('reject negative, fractional, too large and logical operands', () => {
		expect(f('BITAND', -1, 5)).toBe(ERRORS.NUM);
		expect(f('BITOR', 1.5, 5)).toBe(ERRORS.NUM);
		expect(f('BITXOR', 2 ** 48, 1)).toBe(ERRORS.NUM);
		expect(f('BITAND', true, 1)).toBe(ERRORS.VALUE);
		expect(f('BITAND', 'x', 1)).toBe(ERRORS.VALUE);
	});
});

describe('BITLSHIFT / BITRSHIFT', () => {
	it('match the documented examples', () => {
		expect(f('BITLSHIFT', 4, 2)).toBe(16);
		expect(f('BITRSHIFT', 13, 2)).toBe(3);
	});

	it('shift the other way for negative amounts and truncate the amount', () => {
		expect(f('BITLSHIFT', 16, -2)).toBe(4);
		expect(f('BITRSHIFT', 4, -2)).toBe(16);
		expect(f('BITLSHIFT', 1, 3.9)).toBe(8);
		expect(f('BITRSHIFT', MAX, 47)).toBe(1);
		expect(f('BITRSHIFT', 1, 53)).toBe(0);
		expect(f('BITLSHIFT', 1, 47)).toBe(2 ** 47);
	});

	it('validate the range of the result and the shift', () => {
		expect(f('BITLSHIFT', 1, 48)).toBe(ERRORS.NUM);
		expect(f('BITLSHIFT', 2 ** 47, 1)).toBe(ERRORS.NUM);
		expect(f('BITLSHIFT', 1, 54)).toBe(ERRORS.NUM);
		expect(f('BITRSHIFT', 1, -54)).toBe(ERRORS.NUM);
		expect(f('BITRSHIFT', -4, 1)).toBe(ERRORS.NUM);
	});

	it('carry the _xlfn. file prefix', () => {
		for (const name of ['BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT']) {
			expect(fn(ENGINEERING_FUNCTIONS, name).filePrefix).toBe('_xlfn.');
		}
	});
});

import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);

describe('DELTA', () => {
	it('matches the documented examples', () => {
		expect(f('DELTA', 5, 4)).toBe(0);
		expect(f('DELTA', 5, 5)).toBe(1);
		expect(f('DELTA', 0.5, 0)).toBe(0);
		expect(f('DELTA', 0)).toBe(1);
	});

	it('coerces numeric text and rejects other values', () => {
		expect(f('DELTA', '5', 5)).toBe(1);
		expect(f('DELTA', 'five', 5)).toBe(ERRORS.VALUE);
		expect(f('DELTA', true, 1)).toBe(ERRORS.VALUE);
		expect(f('DELTA', arr([1, 2, 3]), 2)).toEqual(arr([0, 1, 0]));
	});
});

describe('GESTEP', () => {
	it('matches the documented examples', () => {
		expect(f('GESTEP', 5, 4)).toBe(1);
		expect(f('GESTEP', 5, 5)).toBe(1);
		expect(f('GESTEP', -4, -5)).toBe(1);
		expect(f('GESTEP', -1)).toBe(0);
		expect(f('GESTEP', 0)).toBe(1);
	});

	it('validates arguments', () => {
		expect(f('GESTEP', 'x')).toBe(ERRORS.VALUE);
		expect(f('GESTEP', 1, false)).toBe(ERRORS.VALUE);
		expect(f('GESTEP', ERRORS.NA)).toBe(ERRORS.NA);
	});
});

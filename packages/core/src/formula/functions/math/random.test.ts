import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import { MATH_FUNCTIONS } from './index';

const RAND = fn(MATH_FUNCTIONS, 'RAND');
const RANDBETWEEN = fn(MATH_FUNCTIONS, 'RANDBETWEEN');

describe('RAND / RANDBETWEEN', () => {
	it('are volatile', () => {
		expect(RAND.volatile).toBe(true);
		expect(RANDBETWEEN.volatile).toBe(true);
	});

	it('draws from the context random source', () => {
		const ctx = new MockContext();
		ctx.randomValues = [0.25, 0];
		expect(call(RAND, [], ctx)).toBe(0.25);
		expect(call(RAND, [], ctx)).toBe(0);
	});

	it('maps the random source onto the integer interval', () => {
		const ctx = new MockContext();
		ctx.randomValues = [0, 0.5, 0.999999, 0.5];
		expect(call(RANDBETWEEN, [1, 10], ctx)).toBe(1);
		expect(call(RANDBETWEEN, [1, 10], ctx)).toBe(6);
		expect(call(RANDBETWEEN, [1, 10], ctx)).toBe(10);
		expect(call(RANDBETWEEN, [-3, -1], ctx)).toBe(-2);
	});

	it('rounds the bounds inward', () => {
		const ctx = new MockContext();
		ctx.randomValues = [0, 0.99];
		expect(call(RANDBETWEEN, [1.5, 2.5], ctx)).toBe(2);
		expect(call(RANDBETWEEN, [1.5, 2.5], ctx)).toBe(2);
	});

	it('validates arguments', () => {
		const ctx = new MockContext();
		expect(call(RANDBETWEEN, [10, 1], ctx)).toBe(ERRORS.NUM);
		expect(call(RANDBETWEEN, [1.2, 1.8], ctx)).toBe(ERRORS.NUM);
		expect(call(RANDBETWEEN, ['a', 1], ctx)).toBe(ERRORS.VALUE);
		expect(call(RANDBETWEEN, [1], ctx)).toBe(ERRORS.VALUE);
		expect(call(RAND, [1], ctx)).toBe(ERRORS.VALUE);
		expect(call(RANDBETWEEN, ['1', '1'], ctx)).toBe(1);
	});
});

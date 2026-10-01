import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { erf, erfc } from './erf';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);

/** [x, erf(x), erfc(x)] from a correctly rounded reference implementation. */
const REFERENCE: Array<[number, number, number]> = [
	[-3, -0.9999779095030015, 1.9999779095030015],
	[-1.5, -0.9661051464753108, 1.9661051464753108],
	[-0.5, -0.5204998778130465, 1.5204998778130465],
	[-0.01, -0.011283415555849616, 1.0112834155558497],
	[0, 0, 1],
	[1e-8, 1.1283791670955126e-8, 0.9999999887162083],
	[0.1, 0.1124629160182849, 0.8875370839817152],
	[0.4999, 0.5204119951613866, 0.47958800483861347],
	[0.5, 0.5204998778130465, 0.4795001221869535],
	[0.745, 0.7079289200957377, 0.2920710799042623],
	[1, 0.8427007929497148, 0.15729920705028516],
	[1.5, 0.9661051464753108, 0.033894853524689274],
	[1.9999, 0.9953201979070294, 0.004679802092970609],
	[2, 0.9953222650189527, 0.0046777349810472645],
	[2.5, 0.999593047982555, 0.000406952017444959],
	[3, 0.9999779095030015, 2.2090496998585438e-5],
	[4.2, 0.9999999971445058, 2.8554941795921843e-9],
	[6, 1, 2.1519736712498913e-17],
	[10, 1, 2.0884875837625446e-45],
	[26.5, 1, 2.2109076642637343e-307],
];

const rel = (got: number, want: number) => Math.abs(got - want) / Math.max(Math.abs(want), 1e-300);

describe('erf / erfc accuracy', () => {
	it.each(REFERENCE)('x = %d', (x, e, ec) => {
		expect(rel(erf(x), e)).toBeLessThan(2e-15);
		expect(rel(erfc(x), ec)).toBeLessThan(5e-15);
	});

	it('handles the far tails', () => {
		expect(erfc(30)).toBe(0);
		expect(erf(30)).toBe(1);
		expect(erfc(-30)).toBe(2);
		expect(erf(-30)).toBe(-1);
	});
});

describe('ERF / ERFC worksheet functions', () => {
	it('match the documented examples', () => {
		expect(f('ERF', 0.745)).toBeCloseTo(0.70792892, 8);
		expect(f('ERF', 1)).toBeCloseTo(0.84270079, 8);
		expect(f('ERF', 0, 1.5)).toBeCloseTo(0.966105146, 9);
		expect(f('ERF.PRECISE', 0.745)).toBeCloseTo(0.70792892, 8);
		expect(f('ERF.PRECISE', 1.5)).toBeCloseTo(0.966105146, 9);
		expect(f('ERFC', 1)).toBeCloseTo(0.15729921, 8);
		expect(f('ERFC.PRECISE', 1)).toBeCloseTo(0.15729921, 8);
	});

	it('integrates between two limits without cancellation', () => {
		expect(rel(f('ERF', 5, 6) as number, 1.5374382746913225e-12)).toBeLessThan(1e-13);
		expect(f('ERF', 1, 0)).toBeCloseTo(-0.8427007929497148, 15);
		expect(f('ERF', -1, 1)).toBeCloseTo(2 * 0.8427007929497148, 15);
		expect(rel(f('ERF', -6, -5) as number, 1.5374382746913225e-12)).toBeLessThan(1e-13);
		expect(f('ERF', 2, 2)).toBe(0);
	});

	it('accepts negative arguments and validates input', () => {
		expect(f('ERF', -0.5)).toBeCloseTo(-0.5204998778130465, 15);
		expect(f('ERFC', -1)).toBeCloseTo(1.8427007929497148, 15);
		expect(f('ERF', 'x')).toBe(ERRORS.VALUE);
		expect(f('ERF', 1, 'x')).toBe(ERRORS.VALUE);
		expect(f('ERFC', true)).toBe(ERRORS.VALUE);
		expect(f('ERF', '0.745')).toBeCloseTo(0.70792892, 8);
		expect(f('ERF', null)).toBe(0);
	});

	it('marks the .PRECISE variants with the _xlfn. prefix', () => {
		expect(fn(ENGINEERING_FUNCTIONS, 'ERF.PRECISE').filePrefix).toBe('_xlfn.');
		expect(fn(ENGINEERING_FUNCTIONS, 'ERFC.PRECISE').filePrefix).toBe('_xlfn.');
		expect(fn(ENGINEERING_FUNCTIONS, 'ERF').filePrefix).toBeUndefined();
	});
});

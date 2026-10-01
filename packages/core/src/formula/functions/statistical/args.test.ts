import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import type { MissingNode } from '../../ast';
import { arr, MockContext } from '../../test-utils/mock-context';
import {
	arrayNumbers,
	coerceArgs,
	define,
	finish,
	isMissing,
	numbersOf,
	scalarFunction,
	valuesOf,
} from './args';
import { rowsOf, stat } from './test-utils/expect';

const MISSING: MissingNode = { type: 'missing' };

describe('argument helpers', () => {
	it('recognises omitted arguments', () => {
		expect(isMissing(undefined)).toBe(true);
		expect(isMissing(MISSING)).toBe(true);
		expect(isMissing(null)).toBe(false);
		expect(isMissing(0)).toBe(false);
		expect(isMissing(ERRORS.NA)).toBe(false);
	});

	it('coerces scalars by spec, applying defaults to omitted arguments', () => {
		const ctx = new MockContext();
		expect(coerceArgs(['n', 'i', 'b'], ['2.5', 3.9, 'TRUE'], ctx)).toEqual([2.5, 3, 1]);
		expect(coerceArgs(['n', { n: 7 }, { b: true }, { i: 3 }], [1], ctx)).toEqual([1, 7, 1, 3]);
		expect(coerceArgs([{ n: 7 }], [MISSING], ctx)).toEqual([7]);
		expect(coerceArgs(['n', 'b'], [null, null], ctx)).toEqual([0, 0]);
		expect(coerceArgs(['n'], ['abc'], ctx)).toBe(ERRORS.VALUE);
		expect(coerceArgs(['b'], ['maybe'], ctx)).toBe(ERRORS.VALUE);
		expect(coerceArgs(['i'], [-2.7], ctx)).toEqual([-2]);
	});

	it('turns non-finite results into #NUM!', () => {
		expect(finish(NaN)).toBe(ERRORS.NUM);
		expect(finish(Infinity)).toBe(ERRORS.NUM);
		expect(finish(ERRORS.DIV0)).toBe(ERRORS.DIV0);
		expect(finish(1.5)).toBe(1.5);
	});

	it('adds the _xlfn. prefix to Excel 2010 names only', () => {
		const impl = () => 0;
		expect(define({ name: 'NORM.DIST', minArgs: 0, maxArgs: 0, params: [], impl }).filePrefix).toBe(
			'_xlfn.',
		);
		expect(
			define({ name: 'NORMDIST', minArgs: 0, maxArgs: 0, params: [], impl }).filePrefix,
		).toBeUndefined();
	});

	it('collects numbers and values from arguments', () => {
		const ctx = new MockContext({ A1: 1, A2: 'x', A3: true, A4: 4 });
		expect(numbersOf([ctx.ref('A1:A4'), '5', MISSING], ctx)).toEqual([1, 4, 5]);
		expect(numbersOf([arr([1, ERRORS.NA])], ctx)).toBe(ERRORS.NA);
		expect(arrayNumbers(ctx.ref('A1:A4'), ctx)).toEqual([1, 4]);
		expect(valuesOf(ctx.ref('A1:A2'), ctx)).toEqual([1, 'x']);
		expect(valuesOf(MISSING, ctx)).toEqual([null]);
	});
});

describe('scalarFunction', () => {
	const DOUBLE = scalarFunction('TEST.DOUBLE', ['n', { n: 1 }], ([x, k]) => x * 2 * k);

	it('propagates argument errors before computing', () => {
		expect(stat('NORM.S.DIST', [ERRORS.NA, true])).toBe(ERRORS.NA);
		expect(stat('NORM.S.DIST', [1, ERRORS.REF])).toBe(ERRORS.REF);
		expect(stat('NORM.S.DIST', ['abc', true])).toBe(ERRORS.VALUE);
	});

	it('enforces the argument count', () => {
		expect(stat('NORM.S.DIST', [1])).toBe(ERRORS.VALUE);
		expect(stat('PHI', [1, 2])).toBe(ERRORS.VALUE);
		expect(DOUBLE.minArgs).toBe(1);
		expect(DOUBLE.maxArgs).toBe(2);
	});

	it('is lifted element-wise over arrays and ranges', () => {
		const ctx = new MockContext({ A1: 0, A2: 1, A3: 'x' });
		const out = rowsOf(stat('NORM.S.DIST', [ctx.ref('A1:A3'), true], ctx));
		expect(out[0][0]).toBe(0.5);
		expect(out[1][0]).toBeCloseTo(0.841344746068543, 15);
		expect(out[2][0]).toBe(ERRORS.VALUE);
		const lifted = rowsOf(stat('PHI', [arr([0, 1])]));
		expect(lifted[0]).toHaveLength(2);
	});

	it('treats blank cells as zero and logical text as logical', () => {
		const ctx = new MockContext({ B1: null });
		expect(stat('NORM.S.DIST', [ctx.ref('B1'), true], ctx)).toBe(0.5);
		expect(stat('NORM.S.DIST', [0, 'FALSE'])).toBeCloseTo(0.3989422804014327, 15);
		expect(stat('NORM.S.DIST', [0, MISSING])).toBeCloseTo(0.3989422804014327, 15);
	});
});

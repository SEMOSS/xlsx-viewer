import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: '1+i', A2: '2-3i', A4: 4, B1: '1+j', B2: 'oops' });
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);

describe('COMPLEX', () => {
	it.each([
		[[3, 4], '3+4i'],
		[[3, 4, 'j'], '3+4j'],
		[[0, 1], 'i'],
		[[1, 0], '1'],
		[[0, 0], '0'],
		[[1.5, -2], '1.5-2i'],
		[[0, -1, 'j'], '-j'],
		[[3, 4, ''], '3+4i'],
		[['3', '4'], '3+4i'],
	])('COMPLEX(%j) = %s', (args, expected) => {
		expect(f('COMPLEX', ...(args as FunctionArg[]))).toBe(expected);
	});

	it('validates the parts and the suffix', () => {
		expect(f('COMPLEX', 3, 4, 'I')).toBe(ERRORS.VALUE);
		expect(f('COMPLEX', 3, 4, 'k')).toBe(ERRORS.VALUE);
		expect(f('COMPLEX', 3, 4, 1)).toBe(ERRORS.VALUE);
		expect(f('COMPLEX', 'a', 4)).toBe(ERRORS.VALUE);
		expect(f('COMPLEX', true, 4)).toBe(ERRORS.VALUE);
	});
});

describe('parts of a complex number', () => {
	it('match the documented examples', () => {
		expect(f('IMABS', '5+12i')).toBe(13);
		expect(f('IMAGINARY', '3+4i')).toBe(4);
		expect(f('IMAGINARY', '0-j')).toBe(-1);
		expect(f('IMAGINARY', 4)).toBe(0);
		expect(f('IMREAL', '6-9i')).toBe(6);
		expect(f('IMARGUMENT', '3+4i')).toBeCloseTo(0.927295218001612, 15);
		expect(f('IMCONJUGATE', '3+4i')).toBe('3-4i');
	});

	it('handle real numbers, signs and suffixes', () => {
		expect(f('IMCONJUGATE', '-2.5j')).toBe('2.5j');
		expect(f('IMCONJUGATE', 5)).toBe('5');
		expect(f('IMABS', -3)).toBe(3);
		expect(f('IMARGUMENT', '-1')).toBeCloseTo(Math.PI, 15);
		expect(f('IMARGUMENT', '-i')).toBeCloseTo(-Math.PI / 2, 15);
		expect(f('IMREAL', 'i')).toBe(0);
	});

	it('validate input', () => {
		expect(f('IMARGUMENT', '0')).toBe(ERRORS.DIV0);
		expect(f('IMABS', '3+4x')).toBe(ERRORS.NUM);
		expect(f('IMREAL', true)).toBe(ERRORS.VALUE);
		expect(f('IMAGINARY', ERRORS.NA)).toBe(ERRORS.NA);
	});
});

describe('arithmetic', () => {
	it('matches the documented examples', () => {
		expect(f('IMSUM', '3+4i', '5-3i')).toBe('8+i');
		expect(f('IMSUB', '13+4i', '5+3i')).toBe('8+i');
		expect(f('IMPRODUCT', '3+4i', '5-3i')).toBe('27+11i');
		expect(f('IMPRODUCT', '1+2i', 30)).toBe('30+60i');
		expect(f('IMDIV', '-238+240i', '10+24i')).toBe('5+12i');
		expect(f('IMPOWER', '2+3i', 3)).toBe('-46+9.00000000000001i');
		expect(f('IMSQRT', '1+i')).toBe('1.09868411346781+0.455089860562227i');
	});

	it('aggregates ranges and arrays, skipping blank cells', () => {
		expect(f('IMSUM', ctx.ref('A1:A4'))).toBe('7-2i');
		expect(f('IMPRODUCT', ctx.ref('A1:A4'))).toBe('20-4i');
		expect(f('IMSUM', arr(['1+i', '2+2i']), '3')).toBe('6+3i');
		expect(f('IMSUM', '1')).toBe('1');
		expect(f('IMPRODUCT', 'i', 'i')).toBe('-1');
	});

	it('keeps the j suffix and rejects mixed suffixes', () => {
		expect(f('IMSUM', '1+j', '2')).toBe('3+j');
		expect(f('IMSUM', '1+j', '1+i')).toBe(ERRORS.VALUE);
		expect(f('IMSUB', '1+j', '1+i')).toBe(ERRORS.VALUE);
		expect(f('IMSUM', ctx.ref('B1:B2'))).toBe(ERRORS.NUM);
	});

	it('handles powers and roots in polar form', () => {
		expect(f('IMPOWER', 'i', 2)).toBe('-1+1.22464679914735E-16i');
		expect(f('IMPOWER', '4', 0.5)).toBe('2');
		expect(f('IMPOWER', '2', -1)).toBe('0.5');
		expect(f('IMPOWER', '0', 2)).toBe('0');
		expect(f('IMPOWER', '0', -1)).toBe(ERRORS.NUM);
		expect(f('IMPOWER', '1+i', 'x')).toBe(ERRORS.VALUE);
		expect(f('IMSQRT', '-4')).toBe('1.22464679914735E-16+2i');
		expect(f('IMSQRT', '0')).toBe('0');
	});

	it('rejects division by zero', () => {
		expect(f('IMDIV', '1+i', '0')).toBe(ERRORS.NUM);
		expect(f('IMDIV', '1+i', 0)).toBe(ERRORS.NUM);
	});

	it('lifts single-value functions over arrays', () => {
		expect(f('IMREAL', arr(['1+2i', '3-4i']))).toEqual(arr([1, 3]));
	});
});

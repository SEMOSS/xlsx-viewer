import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { commonSuffix, formatComplex, parseComplex } from './complex';

describe('parseComplex', () => {
	it.each([
		['3+4i', 3, 4, 'i'],
		['3-4j', 3, -4, 'j'],
		['-2.5j', 0, -2.5, 'j'],
		['i', 0, 1, 'i'],
		['-i', 0, -1, 'i'],
		['+j', 0, 1, 'j'],
		['3+i', 3, 1, 'i'],
		['3-i', 3, -1, 'i'],
		['0i', 0, 0, 'i'],
		['1E3+2.5e-2i', 1000, 0.025, 'i'],
		['.5-.25i', 0.5, -0.25, 'i'],
		['+3-4i', 3, -4, 'i'],
		['1e+2i', 0, 100, 'i'],
	])('%s', (text, re, im, suffix) => {
		expect(parseComplex(text)).toEqual({ re, im, suffix });
	});

	it('reads real numbers, blanks and empty text', () => {
		expect(parseComplex('5')).toEqual({ re: 5, im: 0, suffix: null });
		expect(parseComplex('-1.5e2')).toEqual({ re: -150, im: 0, suffix: null });
		expect(parseComplex(7)).toEqual({ re: 7, im: 0, suffix: null });
		expect(parseComplex(null)).toEqual({ re: 0, im: 0, suffix: null });
		expect(parseComplex('')).toEqual({ re: 0, im: 0, suffix: null });
	});

	it.each(['3+4', '3+4I', '3 + 4i', 'i3', '4ii', '3+4k', 'abc', '1e', '--1', '3+-4i'])(
		'rejects %s',
		(text) => {
			expect(parseComplex(text)).toBe(ERRORS.NUM);
		},
	);

	it('rejects logicals and passes errors through', () => {
		expect(parseComplex(true)).toBe(ERRORS.VALUE);
		expect(parseComplex(ERRORS.NA)).toBe(ERRORS.NA);
	});
});

describe('formatComplex', () => {
	it.each([
		[3, 4, 'i', '3+4i'],
		[3, -4, 'j', '3-4j'],
		[0, 1, 'i', 'i'],
		[0, -1, 'i', '-i'],
		[2, 1, 'j', '2+j'],
		[2, -1, 'i', '2-i'],
		[0, 0, 'i', '0'],
		[-0, -0, 'i', '0'],
		[5, 0, 'i', '5'],
		[0, 2.5, 'i', '2.5i'],
		[1 / 3, -2 / 3, 'i', '0.333333333333333-0.666666666666667i'],
		[1e20, 1e-20, 'i', '1E+20+1E-20i'],
		[-46.00000000000001, 9.000000000000009, 'i', '-46+9.00000000000001i'],
	])('(%d, %d, %s) -> %s', (re, im, suffix, text) => {
		expect(formatComplex(re, im, suffix as 'i' | 'j')).toBe(text);
	});

	it('reports non-finite parts as #NUM!', () => {
		expect(formatComplex(Number.NaN, 0, 'i')).toBe(ERRORS.NUM);
		expect(formatComplex(0, Number.POSITIVE_INFINITY, 'i')).toBe(ERRORS.NUM);
	});
});

describe('commonSuffix', () => {
	it('defaults to i and rejects mixed suffixes', () => {
		expect(commonSuffix([])).toBe('i');
		expect(commonSuffix([{ re: 1, im: 0, suffix: null }])).toBe('i');
		expect(
			commonSuffix([
				{ re: 1, im: 1, suffix: 'j' },
				{ re: 1, im: 0, suffix: null },
			]),
		).toBe('j');
		expect(
			commonSuffix([
				{ re: 1, im: 1, suffix: 'j' },
				{ re: 1, im: 1, suffix: 'i' },
			]),
		).toBe(ERRORS.VALUE);
	});
});

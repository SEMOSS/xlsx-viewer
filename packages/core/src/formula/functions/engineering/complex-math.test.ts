import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);

/** [function, argument, real part, imaginary part] from Python's cmath. */
const REFERENCE: Array<[string, string, number, number]> = [
	['IMEXP', '4+3i', -54.051758861078156, 7.704891372731154],
	['IMEXP', '-0.75-1.5j', 0.033413888102318774, -0.47118326819860185],
	['IMLN', '4+3i', 1.6094379124341003, 0.6435011087932844],
	['IMLN', '-0.75-1.5j', 0.5170368837652692, -2.0344439357957027],
	['IMLN', '0.5', -0.6931471805599453, 0],
	['IMLOG10', '4+3i', 0.6989700043360187, 0.2794689806475475],
	['IMLOG10', '-0.75-1.5j', 0.2245462655597094, -0.8835477750576072],
	['IMLOG2', '4+3i', 2.321928094887362, 0.9283758584626207],
	['IMLOG2', '0.5', -1, 0],
	['IMSIN', '4+3i', -7.619231720321411, -6.5481200409110025],
	['IMSIN', '-0.75-1.5j', -1.6034935732013758, -1.5579700760145927],
	['IMCOS', '4+3i', -6.580663040551157, 7.5815527427465454],
	['IMCOS', '-0.75-1.5j', 1.7212319305052317, -1.451399407513992],
	['IMTAN', '4+3i', 0.00490825806749606, 1.000709536067233],
	['IMTAN', '-0.75-1.5j', -0.09838781913967143, -0.9881121005704792],
	['IMTAN', '0.5', 0.5463024898437905, 0],
	['IMCOT', '4+3i', 0.004901182394304474, -0.9992669278059017],
	['IMCOT', '-0.75-1.5j', -0.09978018312108876, 1.002095657787901],
	['IMSEC', '4+3i', -0.06529402785794704, -0.07522496030277323],
	['IMSEC', '-0.75-1.5j', 0.339547081740266, 0.2863172734172265],
	['IMCSC', '4+3i', -0.0754898329158637, 0.06487747137063549],
	['IMCSC', '-0.75-1.5j', -0.320797366435349, 0.3116898662541724],
	['IMSINH', '4+3i', -27.016813258003936, 3.853738037919377],
	['IMSINH', '-0.75-1.5j', -0.058168384501671194, -1.2914400857052224],
	['IMCOSH', '4+3i', -27.034945603074224, 3.8511533348117775],
	['IMCOSH', '-0.75-1.5j', 0.09158227260398998, 0.8202568175066206],
	['IMSECH', '4+3i', -0.03625349691586887, -0.00516434460775318],
	['IMSECH', '-0.75-1.5j', 0.1344408718164396, -1.2041199516396788],
	['IMCSCH', '4+3i', -0.03627588962862601, -0.005174473184019397],
	['IMCSCH', '-0.75-1.5j', -0.034806335417006486, 0.7727616501491826],
	['IMSQRT', '4+3i', 2.1213203435596424, 0.7071067811865476],
	['IMSQRT', '-0.75-1.5j', 0.6808270643580653, -1.1016013305921617],
];

const close = (got: number, want: number) =>
	Math.abs(got - want) <= 1e-14 * Math.max(1, Math.abs(want));

describe('transcendental complex functions', () => {
	it.each(REFERENCE)('%s(%s)', (name, z, re, im) => {
		const result = f(name, z);
		expect(typeof result).toBe('string');
		expect(close(f('IMREAL', result) as number, re)).toBe(true);
		expect(close(f('IMAGINARY', result) as number, im)).toBe(true);
	});

	it('match the documented examples digit for digit', () => {
		expect(f('IMEXP', '1+i')).toBe('1.46869393991589+2.28735528717884i');
		expect(f('IMLN', '3+4i')).toBe('1.6094379124341+0.927295218001612i');
		expect(f('IMLOG10', '3+4i')).toBe('0.698970004336019+0.402719196273373i');
		expect(f('IMLOG2', '3+4i')).toBe('2.32192809488736+1.33780421245098i');
		expect(f('IMSIN', '3+4i')).toBe('3.85373803791938-27.0168132580039i');
		expect(f('IMCOS', '1+i')).toBe('0.833730025131149-0.988897705762865i');
		expect(f('IMSINH', '-3+4i')).toBe('6.548120040911-7.61923172032141i');
		expect(f('IMTAN', '4+3i')).toBe('0.00490825806749606+1.00070953606723i');
		expect(f('IMSEC', '4+3i')).toBe('-0.0652940278579471-0.0752249603027732i');
		expect(f('IMCSC', '4+3i')).toBe('-0.0754898329158637+0.0648774713706355i');
		expect(f('IMCOSH', '4+3i')).toBe('-27.0349456030742+3.85115333481178i');
	});

	it('keep the suffix of the argument', () => {
		expect(f('IMEXP', '0j')).toBe('1');
		expect(f('IMSIN', '1+j')).toMatch(/j$/);
		expect(f('IMLN', 'j')).toBe('1.5707963267949j');
	});

	it('report poles, logarithms of zero and overflow as #NUM!', () => {
		expect(f('IMLN', '0')).toBe(ERRORS.NUM);
		expect(f('IMLOG10', 0)).toBe(ERRORS.NUM);
		expect(f('IMLOG2', '0i')).toBe(ERRORS.NUM);
		expect(f('IMCOT', '0')).toBe(ERRORS.NUM);
		expect(f('IMCSC', 0)).toBe(ERRORS.NUM);
		expect(f('IMCSCH', 0)).toBe(ERRORS.NUM);
		expect(f('IMEXP', '1000')).toBe(ERRORS.NUM);
		expect(f('IMSIN', 'x')).toBe(ERRORS.NUM);
	});

	it('carry the _xlfn. prefix for the Excel 2013 additions only', () => {
		for (const name of [
			'IMTAN',
			'IMCOT',
			'IMSEC',
			'IMCSC',
			'IMSINH',
			'IMCOSH',
			'IMSECH',
			'IMCSCH',
		]) {
			expect(fn(ENGINEERING_FUNCTIONS, name).filePrefix).toBe('_xlfn.');
		}
		for (const name of ['IMSIN', 'IMCOS', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMSQRT']) {
			expect(fn(ENGINEERING_FUNCTIONS, name).filePrefix).toBeUndefined();
		}
	});
});

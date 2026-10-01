import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(ENGINEERING_FUNCTIONS, name), args, ctx);
const rel = (got: number, want: number) => Math.abs(got - want) / Math.abs(want);

/**
 * [function, x, n, Excel-compatible value, exact value]. The Excel column
 * comes from an independent implementation of the same approximations;
 * the exact column from series / continued-fraction evaluation.
 */
const TABLE: Array<[string, number, number, number, number]> = [
	['BESSELJ', 0.5, 0, 0.9384698074235406, 0.9384698072408129],
	['BESSELY', 0.5, 0, -0.44451873376270784, -0.44451873350670656],
	['BESSELI', 0.5, 0, 1.0634833439946074, 1.0634833707413236],
	['BESSELK', 0.5, 0, 0.9244190350213235, 0.9244190712276659],
	['BESSELJ', 0.5, 1, 0.24226845767957006, 0.24226845767487387],
	['BESSELY', 0.5, 1, -1.4714723918672943, -1.471472392670243],
	['BESSELI', 0.5, 1, 0.25789430328903556, 0.25789430539089625],
	['BESSELK', 0.5, 1, 1.6564411280110791, 1.656441120003301],
	['BESSELJ', 0.5, 2, 0.030604023458698646, 0.030604023458682638],
	['BESSELY', 0.5, 2, -5.441370833706469, -5.441370837174266],
	['BESSELI', 0.5, 2, 0.031906148375295325, 0.03190614917773825],
	['BESSELK', 0.5, 2, 7.5501835470656395, 7.5501835512408695],
	['BESSELJ', 0.5, 5, 8.053627241357472e-6, 8.053627241357475e-6],
	['BESSELY', 0.5, 5, -7946.301473717528, -7946.301478807474],
	['BESSELI', 0.5, 5, 8.22317110629566e-6, 8.223171313109261e-6],
	['BESSELK', 0.5, 5, 12097.979471128536, 12097.979476096392],
	['BESSELJ', 1.5, 0, 0.5118276712499389, 0.5118276717359181],
	['BESSELY', 1.5, 0, 0.38244892476502895, 0.3824489237977589],
	['BESSELI', 1.5, 0, 1.6467232021476756, 1.6467231897728907],
	['BESSELK', 1.5, 0, 0.21380556932365397, 0.21380556264752576],
	['BESSELJ', 1.5, 2, 0.23208767901744795, 0.23208767214421475],
	['BESSELY', 1.5, 2, -0.9321937606262338, -0.9321937597629739],
	['BESSELI', 1.5, 2, 0.33783462087443816, 0.3378346183356807],
	['BESSELK', 1.5, 2, 0.5836559741666656, 0.5836559632566508],
	['BESSELJ', 1.5, 5, 0.0017994217673606115, 0.0017994217673606115],
	['BESSELY', 1.5, 5, -37.190308439288046, -37.19030839549807],
	['BESSELI', 1.5, 5, 0.002170559585408863, 0.0021705595690975563],
	['BESSELK', 1.5, 5, 44.06778193815741, 44.06778115930107],
	['BESSELJ', 2.5, 0, -0.04838377582675685, -0.04838377646819797],
	['BESSELY', 2.5, 0, 0.4980703584466886, 0.498070359615232],
	['BESSELI', 2.5, 0, 3.2898391723912908, 3.289839144050123],
	['BESSELK', 2.5, 0, 0.06234755419101918, 0.06234755320036614],
	['BESSELJ', 2.5, 5, 0.019501625134503528, 0.019501625134503216],
	['BESSELY', 2.5, 5, -3.8301759916745386, -3.830176000740753],
	['BESSELI', 2.5, 5, 0.03284347545496179, 0.03284347517202322],
	['BESSELK', 2.5, 5, 2.7168842921258625, 2.716884290786542],
	['BESSELJ', 5, 1, -0.3275791385663632, -0.32757913759146523],
	['BESSELY', 5, 1, 0.14786313988734273, 0.14786314339122697],
	['BESSELI', 5, 1, 24.335641845705506, 24.33564214245053],
	['BESSELK', 5, 1, 0.004044613383208274, 0.004044613445452165],
	['BESSELJ', 5, 5, 0.2611405469872145, 0.26114054612017007],
	['BESSELY', 5, 5, -0.45369481614111834, -0.453694822491102],
	['BESSELI', 5, 5, 2.157974552930648, 2.1579745473225467],
	['BESSELK', 5, 5, 0.03270627362088703, 0.03270627371203186],
	['BESSELJ', 9.5, 2, 0.2278791540456075, 0.2278791541626919],
	['BESSELY', 9.5, 2, -0.12843591068261181, -0.12843591054225886],
	['BESSELI', 9.5, 2, 1404.3330080268915, 1404.3329741521368],
	['BESSELK', 9.5, 2, 3.671094420744207e-5, 3.671094477065574e-5],
	['BESSELJ', 20, 0, 0.16702466434983618, 0.16702466434058308],
	['BESSELY', 20, 1, -0.16551161434644293, -0.1655116143625212],
	['BESSELI', 20, 2, 39312788.145672314, 39312785.22104077],
	['BESSELK', 20, 5, 1.0538659742087084e-9, 1.0538660139974235e-9],
];

describe('Bessel functions', () => {
	it('match the documented Excel examples', () => {
		expect(f('BESSELI', 1.5, 1)).toBeCloseTo(0.981666428, 9);
		expect(f('BESSELJ', 1.9, 2)).toBeCloseTo(0.329925829, 9);
		expect(f('BESSELK', 1.5, 1)).toBeCloseTo(0.277387804, 9);
		expect(f('BESSELY', 2.5, 1)).toBeCloseTo(0.145918138, 9);
	});

	it.each(TABLE)(
		'%s(%d, %d) reproduces Excel and stays close to the exact value',
		(name, x, n, excel, exact) => {
			const got = f(name, x, n) as number;
			expect(rel(got, excel)).toBeLessThan(1e-13);
			expect(rel(got, exact)).toBeLessThan(5e-7);
		},
	);

	it('uses the parity of the order for negative x', () => {
		expect(f('BESSELJ', -1.9, 2)).toBe(f('BESSELJ', 1.9, 2));
		expect(f('BESSELJ', -1.9, 3)).toBe(-(f('BESSELJ', 1.9, 3) as number));
		expect(f('BESSELJ', -1.9, 1)).toBe(-(f('BESSELJ', 1.9, 1) as number));
		expect(f('BESSELI', -1.5, 1)).toBe(-(f('BESSELI', 1.5, 1) as number));
		expect(f('BESSELI', -1.5, 2)).toBe(f('BESSELI', 1.5, 2));
	});

	it('truncates the order and handles x = 0', () => {
		expect(f('BESSELJ', 1.9, 2.9)).toBe(f('BESSELJ', 1.9, 2));
		expect(f('BESSELJ', 0, 3)).toBe(0);
		expect(f('BESSELI', 0, 3)).toBe(0);
		expect(f('BESSELI', 0, 0)).toBe(1);
	});

	it('validates arguments', () => {
		expect(f('BESSELJ', 1, -1)).toBe(ERRORS.NUM);
		expect(f('BESSELK', 0, 1)).toBe(ERRORS.NUM);
		expect(f('BESSELK', -1, 1)).toBe(ERRORS.NUM);
		expect(f('BESSELY', 0, 1)).toBe(ERRORS.NUM);
		expect(f('BESSELI', 800, 0)).toBe(ERRORS.NUM);
		expect(f('BESSELJ', 'x', 1)).toBe(ERRORS.VALUE);
		expect(f('BESSELJ', true, 1)).toBe(ERRORS.VALUE);
		expect(f('BESSELY', 1, 'n')).toBe(ERRORS.VALUE);
		expect(f('BESSELK', 800, 0)).toBe(0);
	});

	it('lifts over arrays', () => {
		expect(f('BESSELJ', arr([0, 1]), 0)).toEqual(
			arr([f('BESSELJ', 0, 0) as number, f('BESSELJ', 1, 0) as number]),
		);
	});
});

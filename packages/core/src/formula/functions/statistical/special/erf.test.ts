import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { erf, erfc, erfcx, expNegHalfSquare, expNegSquare } from './erf';

// [x, erf(x), erfc(x)] from 60-digit Decimal arithmetic.
const TABLE: [number, number, number][] = [
	[1e-20, 1.1283791670955125e-20, 1.0],
	[0.001, 0.0011283787909692365, 0.9988716212090307],
	[0.1, 0.1124629160182849, 0.887537083981715],
	[0.5, 0.5204998778130465, 0.4795001221869535],
	[0.75, 0.7111556336535151, 0.28884436634648486],
	[1.0, 0.8427007929497149, 0.15729920705028513],
	[1.25, 0.9229001282564583, 0.07709987174354177],
	[1.4999, 0.9660932516622953, 0.03390674833770473],
	[1.5, 0.9661051464753108, 0.033894853524689274],
	[2.0, 0.9953222650189527, 0.004677734981047266],
	[3.0, 0.9999779095030014, 2.209049699858544e-5],
	[5.0, 0.9999999999984626, 1.537459794428035e-12],
	[8.0, 1.0, 1.1224297172982926e-29],
	[15.0, 1.0, 7.212994172451207e-100],
	[26.0, 1.0, 5.663192408856143e-296],
];

describe('erf / erfc', () => {
	it.each(TABLE)('x = %d', (x, e, c) => {
		expectClose(erf(x), e, 2e-15);
		expectClose(erf(-x), -e, 2e-15);
		expectClose(erfc(x), c, 2e-15);
		expectClose(erfc(-x), 2 - c, 2e-15);
	});

	it('handles the special values', () => {
		expect(erf(0)).toBe(0);
		expect(erfc(0)).toBe(1);
		expect(erf(Infinity)).toBe(1);
		expect(erf(-Infinity)).toBe(-1);
		expect(erfc(Infinity)).toBe(0);
		expect(erfc(-Infinity)).toBe(2);
		expect(erf(NaN)).toBeNaN();
		expect(erfc(NaN)).toBeNaN();
		expect(erfc(40)).toBe(0);
	});

	it('scales the complementary function without underflow', () => {
		// erfcx(x) ~ 1 / (x sqrt(pi)) for large x
		expectClose(erfcx(1e10), 1 / (1e10 * Math.sqrt(Math.PI)), 1e-15);
		expectClose(erfcx(30), 0.01879588886141675, 1e-15);
		expectClose(erfcx(0.5) * Math.exp(-0.25), 0.4795001221869535, 1e-15);
		expectClose(erfcx(2) * Math.exp(-4), 0.004677734981047266, 1e-15);
	});

	it('forms exp(-x^2) and exp(-x^2/2) without rounding the square', () => {
		const x = 26.123456789;
		// x^2 rounded to a double would lose about 1e-14 relative in the exponential.
		expectClose(expNegSquare(x), 4.190323809625215e-297, 1e-15);
		expectClose(expNegHalfSquare(37), 5.314068364454539e-298, 1e-15);
		expect(expNegSquare(30)).toBe(0);
		expect(expNegHalfSquare(0)).toBe(1);
	});
});

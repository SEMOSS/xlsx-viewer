import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import {
	factorial,
	gamma,
	lnGamma,
	lnGammaRatio,
	sinPi,
	stirlerr,
	stirlingCorrection,
} from './gamma-function';

// [x, ln Gamma(x), Gamma(x) or null] from 80-digit Decimal arithmetic.
const TABLE: [number, number, number | null][] = [
	[1e-10, 23.025850929882736, 9999999999.422785],
	[0.1, 2.252712651734206, 9.51350769866873],
	[0.5, 0.5723649429247001, 1.772453850905516],
	[0.999999, 5.772164873855652e-7, 1.000000577216654],
	[1.000001, -5.772148423874147e-7, 0.9999994227853242],
	[1.4616321449683622, -0.12148629053584961, 0.8856031944108887],
	[1.99, -0.004195529088791668, 0.9958132598476667],
	[2.00001, 4.22787559764836e-6, 1.0000042278845351],
	[2.5, 0.2846828704729192, 1.329340388179137],
	[4.5, 2.4537365708424423, 11.631728396567448],
	[9.99, 12.779315214350193, 354802.0170198311],
	[10.5, 13.940625219403763, 1133278.3889487856],
	[33.7, 84.00233946014926, 3.0321626547398717e36],
	[100, 359.1342053695754, 9.332621544394415e155],
	[170.5, 704.0044277342047, 5.56209241456e305],
	[1e5, 1051287.7089736569, null],
	[1e100, 2.2925850929940456e102, null],
];

describe('lnGamma and gamma', () => {
	it.each(TABLE)('x = %d', (x, lg, g) => {
		expectClose(lnGamma(x), lg, 1e-15);
		if (g !== null) {
			expectClose(gamma(x), g, 1e-15);
		}
	});

	it('is exact at small integers and zero at 1 and 2', () => {
		expect(lnGamma(1)).toBe(0);
		expect(lnGamma(2)).toBe(0);
		expect(gamma(1)).toBe(1);
		expect(gamma(5)).toBe(24);
		expect(gamma(11)).toBe(3628800);
		expect(gamma(23)).toBe(1124000727777607680000);
		expect(factorial(0)).toBe(1);
		expect(factorial(20)).toBe(2432902008176640000);
		expectClose(factorial(170), 7.257415615307999e306, 1e-15);
	});

	it('uses the reflection formula for negative arguments', () => {
		const neg: [number, number][] = [
			[-0.5, -3.544907701811032],
			[-1.5, 2.363271801207355],
			[-2.5, -0.9453087204829419],
			[-3.75, 0.2678661288614166],
			[-10.1, -2.2134165830856185e-6],
			[-100.3, -1.0431432495783088e-158],
			[-170.2, -2.634834053819688e-307],
		];
		for (const [x, g] of neg) {
			expectClose(gamma(x), g, 1e-15);
		}
	});

	it('returns NaN at the poles and Infinity on overflow', () => {
		expect(gamma(0)).toBeNaN();
		expect(gamma(-1)).toBeNaN();
		expect(gamma(-20)).toBeNaN();
		expect(gamma(172)).toBe(Infinity);
		expect(lnGamma(0)).toBeNaN();
		expect(lnGamma(-2.5)).toBeNaN();
		expect(lnGamma(Infinity)).toBe(Infinity);
	});
});

describe('Stirling helpers', () => {
	it('matches the exact stirlerr table and the series', () => {
		expectClose(stirlerr(1), 0.08106146679532726, 1e-15);
		expectClose(stirlerr(7.5), 0.011104559758206917, 1e-15);
		expectClose(stirlerr(10), stirlingCorrection(10), 1e-15);
		// ln(n!) = (n + 1/2) ln n - n + ln sqrt(2 pi) + stirlerr(n)
		const n = 37.25;
		expectClose(
			(n + 0.5) * Math.log(n) - n + 0.5 * Math.log(2 * Math.PI) + stirlerr(n),
			lnGamma(n + 1),
			1e-15,
		);
	});

	it('forms ln Gamma(h + l) - ln Gamma(h) without cancellation', () => {
		expectClose(lnGammaRatio(1e6, 0.5), 6.907755153982137, 1e-15);
		// ln Gamma(10.5) - ln Gamma(10) from the table above
		expectClose(lnGammaRatio(10, 0.5), 13.940625219403763 - 12.801827480081469, 1e-14);
	});

	it('computes sin(pi x) exactly at integers and half-integers', () => {
		expect(sinPi(3)).toBe(0);
		expect(sinPi(-7)).toBe(0);
		expect(sinPi(0.5)).toBe(1);
		expect(sinPi(1.5)).toBe(-1);
		expect(sinPi(-0.5)).toBe(-1);
		expectClose(sinPi(1e6 + 0.25), Math.SQRT1_2, 1e-15);
		expectClose(sinPi(0.1), Math.sin(Math.PI / 10), 1e-15);
	});
});

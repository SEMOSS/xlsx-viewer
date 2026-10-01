import { describe, expect, it } from 'vitest';

import { expectClose } from '../test-utils/expect';
import { fPdf, fQuantile, fTails } from './f-distribution';
import { tCdf, tPdf, tQuantile, tQuantileTwoTailed, tSf, tTwoTailed } from './student-t';

// [t, df, P(T <= t), density] from 60-digit Decimal arithmetic.
const T_TABLE: [number, number, number, number][] = [
	[-3, 1, 0.10241638234956672, 0.03183098861837907],
	[2, 1, 0.8524163823495667, 0.06366197723675814],
	[60, 1, 0.9946953263673767, 8.839485870141368e-5],
	[0.5, 2, 0.6666666666666666, 0.2962962962962963],
	[8, 3, 0.9979617112061072, 0.0007369065209469263],
	[1.959999998, 60, 0.9726775350120396, 0.059847906435580384],
	[-2.5, 10, 0.015723422118304402, 0.02693872762824446],
	[4, 4, 0.9919349550499538, 0.006708203932499369],
	[10, 30, 0.9999999999771237, 5.327814578274168e-11],
	[1.5, 1000, 0.933034980588957, 0.12950334867626564],
	[-40, 5, 9.205981085886477e-8, 1.1476745739098755e-8],
	[0.1, 7, 0.5384259701976039, 0.3827993342605508],
];

describe('Student t', () => {
	it.each(T_TABLE)('t = %d, df = %d', (t, df, p, f) => {
		expectClose(tCdf(t, df), p, 1e-14);
		expectClose(tSf(-t, df), p, 1e-14);
		expectClose(tPdf(t, df), f, 1e-14);
	});

	it('gives the two-tailed probability', () => {
		expectClose(tTwoTailed(1.959999998, 60), 0.05464492997592096, 1e-13);
		expectClose(tTwoTailed(-1.959999998, 60), 0.05464492997592096, 1e-13);
		expect(tTwoTailed(0, 5)).toBe(1);
		expect(tTwoTailed(1e200, 3)).toBe(0);
		expect(tTwoTailed(NaN, 3)).toBeNaN();
	});

	it.each([
		[0.05, 1, 12.706204736174705],
		[0.01, 3, 5.840909309733357],
		[0.5, 25, 0.6844299649042703],
		[1e-8, 5, 71.69350276982087],
		[0.05, 10, 2.228138851986275],
		[0.05, 1e10, 1.9599639847772814],
	])('two-tailed quantile at p = %d, df = %d', (p, df, t) => {
		expectClose(tQuantileTwoTailed(p, df), t, 1e-13);
	});

	it('inverts the lower tail', () => {
		expectClose(tQuantile(0.999, 30), 3.385184866829305, 1e-14);
		expectClose(tQuantile(0.1, 4), -1.5332062740589438, 1e-14);
		expectClose(tQuantile(0.75, 2), 0.8164965809277261, 1e-14);
		expect(tQuantile(0.5, 9)).toBe(0);
		expect(tQuantileTwoTailed(1, 9)).toBe(0);
	});
});

// [x, d1, d2, P(F <= x), P(F > x), density]
const F_TABLE: [number, number, number, number, number, number][] = [
	[15.2069, 6, 4, 0.9900000430027627, 0.009999956997237313, 0.0012237917087831733],
	[1, 1, 1, 0.5, 0.5, 0.15915494309189535],
	[0.5, 5, 10, 0.22997511934989837, 0.7700248806501017, 0.6876070027706234],
	[3.32583453041301, 5, 10, 0.95, 0.050000000000000086, 0.040586237110367636],
	[2, 10, 20, 0.91021728515625, 0.08978271484375, 0.152740478515625],
	[100, 3, 7, 0.9999958695778987, 4.130422101383355e-6, 1.4090397251052762e-7],
	[0.01, 4, 9, 0.00023979488558792676, 0.9997602051144121, 0.047499788986201025],
	[1.5, 50, 60, 0.9335517053407213, 0.06644829465927873, 0.316733573943473],
];

describe('F distribution', () => {
	it.each(F_TABLE)('x = %d, d1 = %d, d2 = %d', (x, d1, d2, lower, upper, f) => {
		const r = fTails(x, d1, d2);
		expectClose(r.lower, lower, 1e-13);
		expectClose(r.upper, upper, 1e-13);
		expectClose(fPdf(x, d1, d2), f, 1e-13);
	});

	it('handles the boundaries', () => {
		expect(fTails(0, 3, 4)).toEqual({ lower: 0, upper: 1 });
		expect(fTails(Infinity, 3, 4)).toEqual({ lower: 1, upper: 0 });
		expect(fPdf(0, 1, 5)).toBe(Infinity);
		expect(fPdf(0, 2, 5)).toBe(1);
		expect(fPdf(0, 3, 5)).toBe(0);
		expect(fPdf(-1, 3, 5)).toBe(0);
	});

	it.each([
		[0.05, 1, 1, true, 161.44763879758847],
		[0.05, 5, 10, true, 3.325834530413012],
		[0.001, 20, 3, true, 126.41777396613219],
		[0.025, 7, 12, false, 0.21432415255430798],
		[0.9, 2, 2, false, 9],
	])('quantile at %d (d1 = %d, d2 = %d, upper = %s)', (p, d1, d2, upper, x) => {
		expectClose(fQuantile(p, d1, d2, upper), x, 1e-14);
	});
});

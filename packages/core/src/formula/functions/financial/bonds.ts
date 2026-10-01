/**
 * Coupon bonds: PRICE, YIELD, DURATION and MDURATION.
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, readNumbers, REQUIRED } from '../datetime/args';
import {
	couponPeriod,
	daysFromCoupon,
	daysToCoupon,
	periodLength,
	readCouponTerms,
	type CouponTerms,
} from './coupons';
import { bisect, newton } from './solver';

/**
 * Clean price per 100 face value. With one coupon left Excel discounts with
 * simple interest (the inverse of YIELD's closed form); otherwise each cash
 * flow is discounted at the periodic yield.
 */
export function bondPrice(
	t: CouponTerms,
	rate: number,
	yld: number,
	redemption: number,
	date1904: boolean,
): number {
	const p = couponPeriod(t.settlement, t.maturity, t.freq, date1904);
	const e = periodLength(p.pcd, p.ncd, t.freq, t.basis);
	const a = daysFromCoupon(t.settlement, p, t.basis, date1904);
	const dsc = daysToCoupon(t.settlement, p, t.basis, date1904);
	const coupon = (100 * rate) / t.freq;
	const accrued = (coupon * a) / e;
	if (p.num === 1) {
		return (redemption + coupon) / (1 + ((dsc / e) * yld) / t.freq) - accrued;
	}
	const base = 1 + yld / t.freq;
	const x = dsc / e;
	let price = redemption / Math.pow(base, p.num - 1 + x);
	for (let k = 1; k <= p.num; k++) {
		price += coupon / Math.pow(base, k - 1 + x);
	}
	return price - accrued;
}

/**
 * Solve `price(y) = target` for a yield: Newton with a numeric derivative,
 * then bisection on a bracket when Newton fails. Prices fall as yields rise
 * and blow up as the periodic rate approaches -100%.
 */
export function solveYield(
	price: (y: number) => number,
	target: number,
	freq: number,
	guess = 0.05,
): number | null {
	const f = (y: number) => price(y) - target;
	const fdf = (y: number): [number, number] => {
		const h = 1e-7 * Math.max(1, Math.abs(y));
		return [f(y), (f(y + h) - f(y - h)) / (2 * h)];
	};
	const root = newton(fdf, guess, { maxIterations: 100, tolerance: 1e-12 });
	if (root !== null && root > -freq && Math.abs(f(root)) < 1e-9 * Math.max(1, Math.abs(target))) {
		return root;
	}
	let hi = 1;
	while (f(hi) > 0 && hi < 1e6) {
		hi *= 2;
	}
	let lo = 0;
	for (let i = 0; i < 60 && f(lo) < 0; i++) {
		lo = (lo - freq) / 2;
	}
	return bisect(f, lo, hi);
}

function priceImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readCouponTerms(args, ctx, 5, 6);
	const n = readNumbers(args.slice(2), ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(t)) {
		return t;
	}
	if (isCellError(n)) {
		return n;
	}
	const [rate, yld, redemption] = n;
	if (rate < 0 || yld < 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return finite(bondPrice(t, rate, yld, redemption, ctx.date1904));
}

function yieldImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readCouponTerms(args, ctx, 5, 6);
	const n = readNumbers(args.slice(2), ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(t)) {
		return t;
	}
	if (isCellError(n)) {
		return n;
	}
	const [rate, pr, redemption] = n;
	if (rate < 0 || pr <= 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	const p = couponPeriod(t.settlement, t.maturity, t.freq, ctx.date1904);
	if (p.num === 1) {
		const e = periodLength(p.pcd, p.ncd, t.freq, t.basis);
		const a = daysFromCoupon(t.settlement, p, t.basis, ctx.date1904);
		const dsr = daysToCoupon(t.settlement, p, t.basis, ctx.date1904);
		const paid = pr / 100 + ((a / e) * rate) / t.freq;
		return finite(((redemption / 100 + rate / t.freq - paid) / paid) * ((t.freq * e) / dsr));
	}
	const y = solveYield((v) => bondPrice(t, rate, v, redemption, ctx.date1904), pr, t.freq);
	return y === null ? ERRORS.NUM : y;
}

/** Macaulay duration in years (cash flows per 100 face value). */
function macaulay(t: CouponTerms, coupon: number, yld: number, date1904: boolean): number {
	const p = couponPeriod(t.settlement, t.maturity, t.freq, date1904);
	const e = periodLength(p.pcd, p.ncd, t.freq, t.basis);
	const x = (e - daysFromCoupon(t.settlement, p, t.basis, date1904)) / e;
	const base = 1 + yld / t.freq;
	const cash = (100 * coupon) / t.freq;
	let weighted = 0;
	let total = 0;
	for (let k = 1; k <= p.num; k++) {
		const time = k - 1 + x;
		const pv = (k === p.num ? cash + 100 : cash) / Math.pow(base, time);
		weighted += time * pv;
		total += pv;
	}
	return weighted / total / t.freq;
}

function duration(modified: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const t = readCouponTerms(args, ctx, 4, 5);
		const n = readNumbers(args.slice(2), ctx, [REQUIRED, REQUIRED]);
		if (isCellError(t)) {
			return t;
		}
		if (isCellError(n)) {
			return n;
		}
		const [coupon, yld] = n;
		if (coupon < 0 || yld < 0) {
			return ERRORS.NUM;
		}
		const d = macaulay(t, coupon, yld, ctx.date1904);
		return finite(modified ? d / (1 + yld / t.freq) : d);
	};
}

export const BOND_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'PRICE', minArgs: 6, maxArgs: 7, params: ['value'], impl: priceImpl },
	{ name: 'YIELD', minArgs: 6, maxArgs: 7, params: ['value'], impl: yieldImpl },
	{ name: 'DURATION', minArgs: 5, maxArgs: 6, params: ['value'], impl: duration(false) },
	{ name: 'MDURATION', minArgs: 5, maxArgs: 6, params: ['value'], impl: duration(true) },
];

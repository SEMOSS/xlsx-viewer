/**
 * Coupon schedules and the COUP* functions: COUPDAYBS, COUPDAYS,
 * COUPDAYSNC, COUPNCD, COUPNUM and COUPPCD.
 *
 * Coupon dates are counted back from maturity in steps of 12 / frequency
 * months. When maturity is the last day of its month every coupon date is a
 * month end; otherwise the maturity day is kept (clamped to short months).
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { readDates } from '../datetime/args';
import { days360Eu, days360Us, type Basis } from '../datetime/day-count';
import { addMonths, isMonthEnd, serialToYmd } from '../datetime/serial';
import { dayDiff, readBasis, readFrequency, type Frequency } from './securities';

/** The coupon period a settlement date falls in. */
export interface CouponPeriod {
	/** Previous coupon date (on or before settlement). */
	pcd: number;
	/** Next coupon date (after settlement). */
	ncd: number;
	/** Coupons payable between settlement and maturity (COUPNUM). */
	num: number;
}

/** The `k`-th coupon date before (`k > 0`) or after (`k < 0`) an anchor date. */
export function couponDate(anchor: number, k: number, freq: Frequency, date1904: boolean): number {
	const a = serialToYmd(anchor, date1904);
	return addMonths(a, (-k * 12) / freq, isMonthEnd(a, date1904), date1904);
}

/** Coupon period containing `settlement` (requires settlement < maturity). */
export function couponPeriod(
	settlement: number,
	maturity: number,
	freq: Frequency,
	date1904: boolean,
): CouponPeriod {
	const step = 12 / freq;
	const s = serialToYmd(settlement, date1904);
	const m = serialToYmd(maturity, date1904);
	const monthsApart = (m.year - s.year) * 12 + (m.month - s.month);
	let k = Math.max(1, Math.floor(monthsApart / step));
	while (couponDate(maturity, k, freq, date1904) > settlement) {
		k++;
	}
	while (k > 1 && couponDate(maturity, k - 1, freq, date1904) <= settlement) {
		k--;
	}
	return {
		pcd: couponDate(maturity, k, freq, date1904),
		ncd: couponDate(maturity, k - 1, freq, date1904),
		num: k,
	};
}

/** Length of a coupon period in days (COUPDAYS): actual for basis 1, else 360 or 365 / frequency. */
export function periodLength(pcd: number, ncd: number, freq: Frequency, basis: Basis): number {
	if (basis === 1) {
		return ncd - pcd;
	}
	return (basis === 3 ? 365 : 360) / freq;
}

/** Days from the start of the coupon period to settlement (COUPDAYBS). */
export function daysFromCoupon(
	settlement: number,
	p: CouponPeriod,
	basis: Basis,
	date1904: boolean,
): number {
	return dayDiff(p.pcd, settlement, basis, date1904);
}

/** Days from settlement to the next coupon date (COUPDAYSNC). */
export function daysToCoupon(
	settlement: number,
	p: CouponPeriod,
	basis: Basis,
	date1904: boolean,
): number {
	const ymd = (serial: number) => serialToYmd(serial, date1904);
	if (basis === 0) {
		return (
			days360Us(ymd(p.pcd), ymd(p.ncd), true, date1904) -
			days360Us(ymd(p.pcd), ymd(settlement), false, date1904)
		);
	}
	if (basis === 4) {
		return days360Eu(ymd(settlement), ymd(p.ncd));
	}
	return p.ncd - settlement;
}

export interface CouponTerms {
	settlement: number;
	maturity: number;
	freq: Frequency;
	basis: Basis;
}

/**
 * Read `settlement, maturity` (the first two arguments) plus the frequency and basis
 * arguments at the given indexes.
 * Invalid dates are #VALUE!; settlement on or after maturity is #NUM!.
 */
export function readCouponTerms(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	freqIndex: number,
	basisIndex: number,
): CouponTerms | CellError {
	const dates = readDates(args, ctx, 2, ERRORS.VALUE);
	const freq = readFrequency(args[freqIndex], ctx);
	const basis = readBasis(args[basisIndex], ctx);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(freq)) {
		return freq;
	}
	if (isCellError(basis)) {
		return basis;
	}
	if (dates[0] >= dates[1]) {
		return ERRORS.NUM;
	}
	return { settlement: dates[0], maturity: dates[1], freq, basis };
}

function coupon(compute: (t: CouponTerms, p: CouponPeriod, date1904: boolean) => number) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const t = readCouponTerms(args, ctx, 2, 3);
		if (isCellError(t)) {
			return t;
		}
		return compute(t, couponPeriod(t.settlement, t.maturity, t.freq, ctx.date1904), ctx.date1904);
	};
}

const def = (name: string, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: 3,
	maxArgs: 4,
	params: ['value'],
	impl,
});

export const COUPON_FUNCTIONS: FunctionDefinition[] = [
	def(
		'COUPDAYBS',
		coupon((t, p, d) => daysFromCoupon(t.settlement, p, t.basis, d)),
	),
	def(
		'COUPDAYS',
		coupon((t, p) => periodLength(p.pcd, p.ncd, t.freq, t.basis)),
	),
	def(
		'COUPDAYSNC',
		coupon((t, p, d) => daysToCoupon(t.settlement, p, t.basis, d)),
	),
	def(
		'COUPNCD',
		coupon((_t, p) => p.ncd),
	),
	def(
		'COUPNUM',
		coupon((_t, p) => p.num),
	),
	def(
		'COUPPCD',
		coupon((_t, p) => p.pcd),
	),
];

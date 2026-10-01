/**
 * Bonds with an odd first or last period: ODDFPRICE, ODDFYIELD, ODDLPRICE
 * and ODDLYIELD, following the formulas in Excel's documentation
 * (quasi-coupon periods counted back from the first coupon or forward from
 * the last interest date).
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, readDates, readNumbers, REQUIRED } from '../datetime/args';
import type { Basis } from '../datetime/day-count';
import { solveYield } from './bonds';
import { couponDate, couponPeriod, periodLength } from './coupons';
import { dayDiff, readBasis, readFrequency, type Frequency } from './securities';

interface OddFirst {
	settlement: number;
	maturity: number;
	issue: number;
	first: number;
	rate: number;
	redemption: number;
	freq: Frequency;
	basis: Basis;
}

/**
 * Share of the quasi-coupon period `from..to` (normal length `nl`) that lies
 * inside `start..end`: 1 for a complete period, else the basis day count
 * over `nl`, 0 when they do not overlap.
 */
function portion(
	start: number,
	end: number,
	from: number,
	to: number,
	nl: number,
	days: (a: number, b: number) => number,
): number {
	if (start <= from && to <= end) {
		return 1;
	}
	const a = Math.max(start, from);
	const b = Math.min(end, to);
	return b > a ? days(a, b) / nl : 0;
}

/** Clean price of a bond whose first period (issue to first coupon) is short or long. */
export function oddFirstPrice(o: OddFirst, yld: number, date1904: boolean): number {
	const days = (a: number, b: number) => dayDiff(a, b, o.basis, date1904);
	const coupon = (100 * o.rate) / o.freq;
	const base = 1 + yld / o.freq;
	// Regular coupons after the first one.
	const n = couponPeriod(o.first, o.maturity, o.freq, date1904).num;
	const quasiStart = couponDate(o.first, 1, o.freq, date1904);
	if (o.issue >= quasiStart) {
		// Short first period: the odd coupon is pro-rated over a normal period.
		const e = periodLength(quasiStart, o.first, o.freq, o.basis);
		const x = days(o.settlement, o.first) / e;
		let price =
			o.redemption / Math.pow(base, n + x) +
			(coupon * days(o.issue, o.first)) / e / Math.pow(base, x);
		for (let k = 1; k <= n; k++) {
			price += coupon / Math.pow(base, k + x);
		}
		return price - (coupon * days(o.issue, o.settlement)) / e;
	}
	// Long first period: sum over the quasi-coupon periods it spans.
	let nc = 1;
	while (couponDate(o.first, nc, o.freq, date1904) > o.issue) {
		nc++;
	}
	let dcOverNl = 0;
	let aOverNl = 0;
	for (let j = nc; j >= 1; j--) {
		const from = couponDate(o.first, j, o.freq, date1904);
		const to = couponDate(o.first, j - 1, o.freq, date1904);
		const nl = periodLength(from, to, o.freq, o.basis);
		dcOverNl += j === nc ? days(o.issue, to) / nl : 1;
		aOverNl += portion(o.issue, o.settlement, from, to, nl, days);
	}
	// Quasi-coupon period holding the settlement date.
	let js = 1;
	while (couponDate(o.first, js, o.freq, date1904) > o.settlement) {
		js++;
	}
	const qs = couponDate(o.first, js, o.freq, date1904);
	const qn = couponDate(o.first, js - 1, o.freq, date1904);
	const nq = js - 1;
	const e = periodLength(qs, qn, o.freq, o.basis);
	const dsc = o.basis === 2 || o.basis === 3 ? qn - o.settlement : e - days(qs, o.settlement);
	const x = dsc / e;
	let price =
		o.redemption / Math.pow(base, n + nq + x) + (coupon * dcOverNl) / Math.pow(base, nq + x);
	for (let k = 1; k <= n; k++) {
		price += coupon / Math.pow(base, k + nq + x);
	}
	return price - coupon * aOverNl;
}

/** `settlement, maturity, issue, first_coupon, rate, yld|pr, redemption, frequency, [basis]`. */
function readOddFirst(
	args: FunctionArg[],
	ctx: FunctionContext,
): { o: OddFirst; value: number } | CellError {
	const dates = readDates(args, ctx, 4, ERRORS.VALUE);
	const n = readNumbers(args.slice(4), ctx, [REQUIRED, REQUIRED, REQUIRED]);
	const freq = readFrequency(args[7], ctx);
	const basis = readBasis(args[8], ctx);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(freq)) {
		return freq;
	}
	if (isCellError(basis)) {
		return basis;
	}
	const [settlement, maturity, issue, first] = dates;
	const [rate, value, redemption] = n;
	if (!(maturity > first && first > settlement && settlement > issue)) {
		return ERRORS.NUM;
	}
	if (rate < 0 || value < 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return { o: { settlement, maturity, issue, first, rate, redemption, freq, basis }, value };
}

function oddfPriceImpl(args: FunctionArg[], ctx: FunctionContext) {
	const r = readOddFirst(args, ctx);
	return isCellError(r) ? r : finite(oddFirstPrice(r.o, r.value, ctx.date1904));
}

function oddfYieldImpl(args: FunctionArg[], ctx: FunctionContext) {
	const r = readOddFirst(args, ctx);
	if (isCellError(r)) {
		return r;
	}
	if (r.value <= 0) {
		return ERRORS.NUM;
	}
	const y = solveYield((v) => oddFirstPrice(r.o, v, ctx.date1904), r.value, r.o.freq);
	return y === null ? ERRORS.NUM : y;
}

interface OddLastSums {
	/** Sum of DC_i / NL_i: coupon days of each quasi-coupon period. */
	dc: number;
	/** Sum of A_i / NL_i: accrued days up to settlement. */
	a: number;
	/** Sum of DSC_i / NL_i: days from settlement to maturity. */
	dsc: number;
}

/** Quasi-coupon sums of an odd last period running from the last interest date to maturity. */
function oddLastSums(
	settlement: number,
	maturity: number,
	last: number,
	freq: Frequency,
	basis: Basis,
	date1904: boolean,
): OddLastSums {
	const days = (a: number, b: number) => dayDiff(a, b, basis, date1904);
	const sums: OddLastSums = { dc: 0, a: 0, dsc: 0 };
	for (let j = 1; ; j++) {
		const from = couponDate(last, -(j - 1), freq, date1904);
		const to = couponDate(last, -j, freq, date1904);
		const nl = periodLength(from, to, freq, basis);
		sums.dc += portion(from, maturity, from, to, nl, days);
		sums.a += portion(from, Math.min(settlement, maturity), from, to, nl, days);
		sums.dsc += portion(settlement, maturity, from, to, nl, days);
		if (to >= maturity) {
			return sums;
		}
	}
}

/** `settlement, maturity, last_interest, rate, yld|pr, redemption, frequency, [basis]`. */
function oddLast(solveForYield: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const dates = readDates(args, ctx, 3, ERRORS.VALUE);
		const n = readNumbers(args.slice(3), ctx, [REQUIRED, REQUIRED, REQUIRED]);
		const freq = readFrequency(args[6], ctx);
		const basis = readBasis(args[7], ctx);
		if (isCellError(dates)) {
			return dates;
		}
		if (isCellError(n)) {
			return n;
		}
		if (isCellError(freq)) {
			return freq;
		}
		if (isCellError(basis)) {
			return basis;
		}
		const [settlement, maturity, last] = dates;
		const [rate, value, redemption] = n;
		if (!(maturity > settlement && settlement > last)) {
			return ERRORS.NUM;
		}
		if (rate < 0 || redemption <= 0 || (solveForYield ? value <= 0 : value < 0)) {
			return ERRORS.NUM;
		}
		const s = oddLastSums(settlement, maturity, last, freq, basis, ctx.date1904);
		const coupon = (100 * rate) / freq;
		const total = redemption + s.dc * coupon;
		if (!solveForYield) {
			return finite(total / (1 + (s.dsc * value) / freq) - s.a * coupon);
		}
		const paid = value + s.a * coupon;
		return finite(((total - paid) / paid) * (freq / s.dsc));
	};
}

const def = (name: string, args: number, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: args - 1,
	maxArgs: args,
	params: ['value'],
	impl,
});

export const ODD_COUPON_FUNCTIONS: FunctionDefinition[] = [
	def('ODDFPRICE', 9, oddfPriceImpl),
	def('ODDFYIELD', 9, oddfYieldImpl),
	def('ODDLPRICE', 8, oddLast(false)),
	def('ODDLYIELD', 8, oddLast(true)),
];

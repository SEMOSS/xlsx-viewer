/**
 * Accrued interest: ACCRINT (periodic interest) and ACCRINTM (interest at maturity).
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { boolArg, finite, readDates, readNumbers, REQUIRED } from '../datetime/args';
import { yearFrac, type Basis } from '../datetime/day-count';
import { serialToYmd } from '../datetime/serial';
import { couponDate, periodLength } from './coupons';
import { dayDiff, readBasis, readFrequency, type Frequency } from './securities';

/** Index `j` of the quasi-coupon date `couponDate(anchor, -j)` on or before `date`. */
function periodIndex(anchor: number, date: number, freq: Frequency, date1904: boolean): number {
	const a = serialToYmd(anchor, date1904);
	const d = serialToYmd(date, date1904);
	let j = Math.floor(((d.year - a.year) * 12 + (d.month - a.month)) / (12 / freq));
	while (couponDate(anchor, -j, freq, date1904) > date) {
		j--;
	}
	while (couponDate(anchor, -(j + 1), freq, date1904) <= date) {
		j++;
	}
	return j;
}

/**
 * Accrued interest from `start` to `settlement`, summed over the
 * quasi-coupon periods of the schedule anchored at the first interest date:
 * `par * rate / freq * sum(A_i / NL_i)`. A complete period counts as exactly
 * one coupon; partial periods use the basis day count over the normal
 * period length.
 */
export function accruedInterest(
	start: number,
	settlement: number,
	firstInterest: number,
	rate: number,
	par: number,
	freq: Frequency,
	basis: Basis,
	date1904: boolean,
): number {
	let sum = 0;
	for (let j = periodIndex(firstInterest, start, freq, date1904); ; j++) {
		const from = couponDate(firstInterest, -j, freq, date1904);
		if (from >= settlement) {
			break;
		}
		const to = couponDate(firstInterest, -(j + 1), freq, date1904);
		if (start <= from && to <= settlement) {
			sum += 1;
			continue;
		}
		const a = Math.max(start, from);
		const b = Math.min(settlement, to);
		if (b > a) {
			sum += dayDiff(a, b, basis, date1904) / periodLength(from, to, freq, basis);
		}
	}
	return ((par * rate) / freq) * sum;
}

function accrintImpl(args: FunctionArg[], ctx: FunctionContext) {
	const dates = readDates(args, ctx, 3, ERRORS.VALUE);
	const n = readNumbers(args.slice(3), ctx, [REQUIRED, 1000]);
	const freq = readFrequency(args[5], ctx);
	const basis = readBasis(args[6], ctx);
	const fromIssue = boolArg(args[7], true);
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
	if (isCellError(fromIssue)) {
		return fromIssue;
	}
	const [issue, firstInterest, settlement] = dates;
	const [rate, par] = n;
	if (rate <= 0 || par <= 0 || issue >= settlement) {
		return ERRORS.NUM;
	}
	// calc_method FALSE: only the interest accrued since the first interest date.
	const start = !fromIssue && settlement > firstInterest ? firstInterest : issue;
	return finite(
		accruedInterest(start, settlement, firstInterest, rate, par, freq, basis, ctx.date1904),
	);
}

function accrintmImpl(args: FunctionArg[], ctx: FunctionContext) {
	const dates = readDates(args, ctx, 2, ERRORS.VALUE);
	const n = readNumbers(args.slice(2), ctx, [REQUIRED, 1000]);
	const basis = readBasis(args[4], ctx);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(basis)) {
		return basis;
	}
	const [issue, settlement] = dates;
	const [rate, par] = n;
	if (rate <= 0 || par <= 0 || issue >= settlement) {
		return ERRORS.NUM;
	}
	return finite(par * rate * yearFrac(issue, settlement, basis, ctx.date1904));
}

export const ACCRUED_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'ACCRINT', minArgs: 6, maxArgs: 8, params: ['value'], impl: accrintImpl },
	{ name: 'ACCRINTM', minArgs: 3, maxArgs: 5, params: ['value'], impl: accrintmImpl },
];

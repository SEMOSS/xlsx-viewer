/**
 * Discounted and interest-at-maturity securities: DISC, INTRATE, RECEIVED,
 * PRICEDISC, PRICEMAT, YIELDDISC and YIELDMAT.
 *
 * Day ratios `days / B` use the basis day count over the year basis, which
 * for actual/actual is YEARFRAC's year length of the span.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, readDates, readNumbers, REQUIRED } from '../datetime/args';
import { yearFrac, type Basis } from '../datetime/day-count';
import { readBasis } from './securities';

interface Terms {
	settlement: number;
	maturity: number;
	basis: Basis;
	/** The numeric arguments after the two dates. */
	values: number[];
}

/**
 * Read `settlement, maturity` (invalid dates are #VALUE!), the required
 * numbers in `numbers` and the optional basis. Settlement on or after
 * maturity is #NUM!.
 */
function readTerms(
	ctx: FunctionContext,
	dateArgs: FunctionArg[],
	numbers: FunctionArg[],
	basisArg: FunctionArg | undefined,
): Terms | CellError {
	const dates = readDates(dateArgs, ctx, 2, ERRORS.VALUE);
	const values = readNumbers(
		numbers,
		ctx,
		numbers.map(() => REQUIRED),
	);
	const basis = readBasis(basisArg, ctx);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(values)) {
		return values;
	}
	if (isCellError(basis)) {
		return basis;
	}
	if (dates[0] >= dates[1]) {
		return ERRORS.NUM;
	}
	return { settlement: dates[0], maturity: dates[1], basis, values };
}

/** `settlement, maturity, x, y, [basis]` layout shared by DISC, INTRATE, ... */
function readSimple(args: FunctionArg[], ctx: FunctionContext): Terms | CellError {
	return readTerms(ctx, [args[0], args[1]], [args[2], args[3]], args[4]);
}

/** `settlement, maturity, issue, rate, x, [basis]` layout of PRICEMAT / YIELDMAT. */
function readWithIssue(
	args: FunctionArg[],
	ctx: FunctionContext,
): (Terms & { issue: number }) | CellError {
	const t = readTerms(ctx, [args[0], args[1]], [args[3], args[4]], args[5]);
	if (isCellError(t)) {
		return t;
	}
	const issue = readDates([args[2]], ctx, 1, ERRORS.VALUE);
	if (isCellError(issue)) {
		return issue;
	}
	return issue[0] > t.settlement ? ERRORS.NUM : { ...t, issue: issue[0] };
}

/** Fraction of a year from settlement to maturity (DSM / B). */
function term(t: Terms, ctx: FunctionContext): number {
	return yearFrac(t.settlement, t.maturity, t.basis, ctx.date1904);
}

function discImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readSimple(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [pr, redemption] = t.values;
	if (pr <= 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return finite((redemption - pr) / redemption / term(t, ctx));
}

function intrateImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readSimple(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [investment, redemption] = t.values;
	if (investment <= 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return finite((redemption - investment) / investment / term(t, ctx));
}

function receivedImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readSimple(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [investment, discount] = t.values;
	if (investment <= 0 || discount <= 0) {
		return ERRORS.NUM;
	}
	const factor = 1 - discount * term(t, ctx);
	return factor <= 0 ? ERRORS.NUM : finite(investment / factor);
}

function priceDiscImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readSimple(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [discount, redemption] = t.values;
	if (discount <= 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return finite(redemption - discount * redemption * term(t, ctx));
}

function yieldDiscImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readSimple(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [pr, redemption] = t.values;
	if (pr <= 0 || redemption <= 0) {
		return ERRORS.NUM;
	}
	return finite((redemption - pr) / pr / term(t, ctx));
}

/** Year fractions issue-to-maturity, issue-to-settlement and settlement-to-maturity. */
function maturityFractions(t: Terms, issue: number, ctx: FunctionContext) {
	return {
		dim: yearFrac(issue, t.maturity, t.basis, ctx.date1904),
		a: yearFrac(issue, t.settlement, t.basis, ctx.date1904),
		dsm: term(t, ctx),
	};
}

function priceMatImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readWithIssue(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [rate, yld] = t.values;
	if (rate < 0 || yld < 0) {
		return ERRORS.NUM;
	}
	const f = maturityFractions(t, t.issue, ctx);
	return finite(((1 + f.dim * rate) / (1 + f.dsm * yld) - f.a * rate) * 100);
}

function yieldMatImpl(args: FunctionArg[], ctx: FunctionContext) {
	const t = readWithIssue(args, ctx);
	if (isCellError(t)) {
		return t;
	}
	const [rate, pr] = t.values;
	if (rate < 0 || pr <= 0) {
		return ERRORS.NUM;
	}
	const f = maturityFractions(t, t.issue, ctx);
	const paid = pr / 100 + f.a * rate;
	return finite((1 + f.dim * rate - paid) / paid / f.dsm);
}

const def = (
	name: string,
	min: number,
	max: number,
	impl: FunctionDefinition['impl'],
): FunctionDefinition => ({
	name,
	minArgs: min,
	maxArgs: max,
	params: ['value'],
	impl,
});

export const DISCOUNT_FUNCTIONS: FunctionDefinition[] = [
	def('DISC', 4, 5, discImpl),
	def('INTRATE', 4, 5, intrateImpl),
	def('RECEIVED', 4, 5, receivedImpl),
	def('PRICEDISC', 4, 5, priceDiscImpl),
	def('PRICEMAT', 5, 6, priceMatImpl),
	def('YIELDDISC', 4, 5, yieldDiscImpl),
	def('YIELDMAT', 5, 6, yieldMatImpl),
];

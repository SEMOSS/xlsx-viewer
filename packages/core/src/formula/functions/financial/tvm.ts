/**
 * Time value of money: PV, FV, PMT, IPMT, PPMT, NPER, ISPMT, CUMIPMT and CUMPRINC.
 *
 * Sign convention (as in Excel): cash paid out is negative, cash received
 * positive. `type` is 0 for payments at the end of each period and 1 for
 * payments at the beginning (any non-zero value counts as 1).
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, readNumbers, REQUIRED } from '../datetime/args';

/** Payment type flag: 1 when payments are due at the beginning of periods. */
export function payType(type: number): 0 | 1 {
	return type !== 0 ? 1 : 0;
}

export function pmtOf(rate: number, nper: number, pv: number, fv: number, type: 0 | 1): number {
	if (rate === 0) {
		return -(pv + fv) / nper;
	}
	const f = Math.pow(1 + rate, nper);
	return (-rate * (pv * f + fv)) / ((1 + rate * type) * (f - 1));
}

export function fvOf(rate: number, nper: number, pmt: number, pv: number, type: 0 | 1): number {
	if (rate === 0) {
		return -(pv + pmt * nper);
	}
	const f = Math.pow(1 + rate, nper);
	return -(pv * f + (pmt * (1 + rate * type) * (f - 1)) / rate);
}

export function pvOf(rate: number, nper: number, pmt: number, fv: number, type: 0 | 1): number {
	if (rate === 0) {
		return -(fv + pmt * nper);
	}
	const f = Math.pow(1 + rate, nper);
	return -(fv + (pmt * (1 + rate * type) * (f - 1)) / rate) / f;
}

/** Interest part of payment `per` (LibreOffice / Excel formulation). */
export function ipmtOf(
	rate: number,
	per: number,
	nper: number,
	pv: number,
	fv: number,
	type: 0 | 1,
): number {
	const pmt = pmtOf(rate, nper, pv, fv, type);
	let interest: number;
	if (per === 1) {
		interest = type === 1 ? 0 : -pv;
	} else if (type === 1) {
		interest = fvOf(rate, per - 2, pmt, pv, 1) - pmt;
	} else {
		interest = fvOf(rate, per - 1, pmt, pv, 0);
	}
	return interest * rate;
}

function pvImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, 0, 0]);
	return isCellError(n) ? n : finite(pvOf(n[0], n[1], n[2], n[3], payType(n[4])));
}

function fvImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, 0, 0]);
	return isCellError(n) ? n : finite(fvOf(n[0], n[1], n[2], n[3], payType(n[4])));
}

function pmtImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, 0, 0]);
	return isCellError(n) ? n : finite(pmtOf(n[0], n[1], n[2], n[3], payType(n[4])));
}

function periodPayment(principal: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED, 0, 0]);
		if (isCellError(n)) {
			return n;
		}
		const [rate, per, nper, pv, fv] = n;
		const type = payType(n[5]);
		if (per < 1 || per > nper) {
			return ERRORS.NUM;
		}
		const interest = ipmtOf(rate, per, nper, pv, fv, type);
		return finite(principal ? pmtOf(rate, nper, pv, fv, type) - interest : interest);
	};
}

function nperImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, 0, 0]);
	if (isCellError(n)) {
		return n;
	}
	const [rate, pmt, pv, fv] = n;
	const type = payType(n[4]);
	if (rate === 0) {
		return finite(-(pv + fv) / pmt);
	}
	const a = pmt * (1 + rate * type);
	return finite(Math.log((a - fv * rate) / (pv * rate + a)) / Math.log(1 + rate));
}

function ispmtImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [rate, per, nper, pv] = n;
	return nper === 0 ? ERRORS.DIV0 : pv * rate * (per / nper - 1);
}

/** CUMIPMT / CUMPRINC: interest or principal paid between two periods (inclusive). */
function cumulative(principal: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED, REQUIRED, REQUIRED]);
		if (isCellError(n)) {
			return n;
		}
		const rate = n[0];
		const pv = n[2];
		const [nper, startPer, endPer, typeArg] = [n[1], n[3], n[4], n[5]].map(Math.trunc);
		if (rate <= 0 || nper <= 0 || pv <= 0 || startPer < 1 || endPer < startPer || endPer > nper) {
			return ERRORS.NUM;
		}
		if (typeArg !== 0 && typeArg !== 1) {
			return ERRORS.NUM;
		}
		const type: 0 | 1 = typeArg === 1 ? 1 : 0;
		const pmt = pmtOf(rate, nper, pv, 0, type);
		let interest = 0;
		let principalPaid = 0;
		for (let per = startPer; per <= endPer; per++) {
			const i = ipmtOf(rate, per, nper, pv, 0, type);
			interest += i;
			principalPaid += pmt - i;
		}
		return finite(principal ? principalPaid : interest);
	};
}

const value = (
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

export const TVM_FUNCTIONS: FunctionDefinition[] = [
	value('PV', 3, 5, pvImpl),
	value('FV', 3, 5, fvImpl),
	value('PMT', 3, 5, pmtImpl),
	value('IPMT', 4, 6, periodPayment(false)),
	value('PPMT', 4, 6, periodPayment(true)),
	value('NPER', 3, 5, nperImpl),
	value('ISPMT', 4, 4, ispmtImpl),
	value('CUMIPMT', 6, 6, cumulative(false)),
	value('CUMPRINC', 6, 6, cumulative(true)),
];

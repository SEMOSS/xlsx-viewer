/**
 * Interest rate functions: RATE, EFFECT, NOMINAL, PDURATION, RRI, FVSCHEDULE.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { forEachScalar } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, numArg, readNumbers, REQUIRED } from '../datetime/args';
import { newton } from './solver';
import { payType } from './tvm';

/** Excel documents RATE as converging to 1e-7 within 20 iterations, else #NUM!. */
const RATE_ITERATIONS = 20;
const RATE_TOLERANCE = 1e-7;

/**
 * Solve `pv*(1+r)^n + pmt*(1+r*type)*((1+r)^n-1)/r + fv = 0` for `r` with
 * Newton's method from `guess`, the way Excel does (no fallback search).
 */
export function solveRate(
	nper: number,
	pmt: number,
	pv: number,
	fv: number,
	type: 0 | 1,
	guess: number,
): number | null {
	// Payments in advance are equivalent to end-of-period payments with shifted values.
	const p0 = type === 1 ? pv + pmt : pv;
	const f0 = type === 1 ? fv - pmt : fv;
	const fdf = (x: number): [number, number] => {
		if (x <= -1) {
			return [Number.NaN, Number.NaN];
		}
		const lg = Math.log1p(x);
		const pn = Math.exp(nper * lg);
		const pn1 = pn / (1 + x);
		let g: number;
		let dg: number;
		if (Math.abs(x) < 1e-9) {
			g = nper + ((nper * (nper - 1)) / 2) * x;
			dg = (nper * (nper - 1)) / 2 + ((nper * (nper - 1) * (nper - 2)) / 6) * x;
		} else {
			g = Math.expm1(nper * lg) / x;
			dg = (nper * pn1 - g) / x;
		}
		return [f0 + p0 * pn + pmt * g, p0 * nper * pn1 + pmt * dg];
	};
	const root = newton(fdf, guess, {
		maxIterations: RATE_ITERATIONS,
		tolerance: RATE_TOLERANCE,
		polish: 6,
	});
	return root !== null && root > -1 ? root : null;
}

function rateImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, 0, 0, 0.1]);
	if (isCellError(n)) {
		return n;
	}
	const [nper, pmt, pv, fv, type, guess] = n;
	if (nper <= 0) {
		return ERRORS.NUM;
	}
	const r = solveRate(nper, pmt, pv, fv, payType(type), guess);
	return r === null ? ERRORS.NUM : r;
}

function effectImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const nominal = n[0];
	const npery = Math.trunc(n[1]);
	if (nominal <= 0 || npery < 1) {
		return ERRORS.NUM;
	}
	return finite(Math.pow(1 + nominal / npery, npery) - 1);
}

function nominalImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const effect = n[0];
	const npery = Math.trunc(n[1]);
	if (effect <= 0 || npery < 1) {
		return ERRORS.NUM;
	}
	return finite(npery * (Math.pow(1 + effect, 1 / npery) - 1));
}

function pdurationImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [rate, pv, fv] = n;
	if (rate <= 0 || pv <= 0 || fv <= 0) {
		return ERRORS.NUM;
	}
	return finite((Math.log(fv) - Math.log(pv)) / Math.log1p(rate));
}

function rriImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [nper, pv, fv] = n;
	if (nper <= 0 || pv === 0) {
		return ERRORS.NUM;
	}
	return finite(Math.pow(fv / pv, 1 / nper) - 1);
}

/** FVSCHEDULE: blanks in the schedule count as 0%; anything but numbers is #VALUE!. */
function fvScheduleImpl(args: FunctionArg[], ctx: FunctionContext) {
	const principal = numArg(args[0], ctx);
	if (isCellError(principal)) {
		return principal;
	}
	let result = principal;
	let error: CellError | undefined;
	forEachScalar(args[1], ctx, (v, fromRef) => {
		if (error || v === null) {
			return;
		}
		if (isCellError(v)) {
			error = v;
		} else if (typeof v === 'number') {
			result *= 1 + v;
		} else if (!fromRef && typeof v === 'string') {
			const rate = numArg(v, ctx);
			if (isCellError(rate)) {
				error = rate;
			} else {
				result *= 1 + rate;
			}
		} else {
			error = ERRORS.VALUE;
		}
	});
	return error ?? finite(result);
}

export const RATE_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'RATE', minArgs: 3, maxArgs: 6, params: ['value'], impl: rateImpl },
	{ name: 'EFFECT', minArgs: 2, maxArgs: 2, params: ['value'], impl: effectImpl },
	{ name: 'NOMINAL', minArgs: 2, maxArgs: 2, params: ['value'], impl: nominalImpl },
	{
		name: 'PDURATION',
		minArgs: 3,
		maxArgs: 3,
		params: ['value'],
		filePrefix: '_xlfn.',
		impl: pdurationImpl,
	},
	{ name: 'RRI', minArgs: 3, maxArgs: 3, params: ['value'], filePrefix: '_xlfn.', impl: rriImpl },
	{ name: 'FVSCHEDULE', minArgs: 2, maxArgs: 2, params: ['value', 'any'], impl: fvScheduleImpl },
];

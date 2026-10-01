/**
 * Shared pieces of the securities functions: frequency / basis arguments
 * and basis-dependent day counting.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext } from '../../types';
import { intArg } from '../datetime/args';
import { days360Eu, days360Us, isBasis, type Basis } from '../datetime/day-count';
import { serialToYmd } from '../datetime/serial';

export type Frequency = 1 | 2 | 4;

/** Coupon frequency (annual, semi-annual or quarterly); truncated, else #NUM!. */
export function readFrequency(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
): Frequency | CellError {
	const n = intArg(arg, ctx);
	if (isCellError(n)) {
		return n;
	}
	return n === 1 || n === 2 || n === 4 ? n : ERRORS.NUM;
}

/** Day-count basis 0-4 (default 0); truncated, else #NUM!. */
export function readBasis(arg: FunctionArg | undefined, ctx: FunctionContext): Basis | CellError {
	const n = intArg(arg, ctx, 0);
	if (isCellError(n)) {
		return n;
	}
	return isBasis(n) ? n : ERRORS.NUM;
}

/** Days from `start` to `end` counted with a basis (30/360 for bases 0 and 4, actual otherwise). */
export function dayDiff(start: number, end: number, basis: Basis, date1904: boolean): number {
	if (basis === 0) {
		return days360Us(serialToYmd(start, date1904), serialToYmd(end, date1904), false, date1904);
	}
	if (basis === 4) {
		return days360Eu(serialToYmd(start, date1904), serialToYmd(end, date1904));
	}
	return end - start;
}

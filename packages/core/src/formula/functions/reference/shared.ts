import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../../address/constants';
import type { RangeAddress } from '../../../types/address';
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type { FunctionArg, FunctionContext, RefValue, Scalar } from '../../types';
import { isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from '../logical/lift';

/** Helpers shared by the reference functions. */

/** The argument was not supplied, written empty, or is blank. */
export function isOmittedArg(args: readonly FunctionArg[], i: number): boolean {
	const a = args[i];
	return a === undefined || a === null || isMissingNode(a);
}

/**
 * A single-area reference argument. Errors pass through; unions of several
 * areas are `#REF!` and anything that is not a reference `#VALUE!`.
 */
export function singleArea(arg: FunctionArg | undefined): RefValue | CellError {
	if (isRefValue(arg)) {
		return arg;
	}
	if (isMultiRefValue(arg)) {
		return arg.refs.length === 1 ? arg.refs[0] : ERRORS.REF;
	}
	if (isCellError(arg)) {
		return arg;
	}
	return ERRORS.VALUE;
}

/** Integer value of a scalar argument (truncated toward zero; blank / empty is `fallback`). */
export function integerOf(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
	fallback = 0,
): number | CellError {
	if (arg === undefined || arg === null || isMissingNode(arg)) {
		return fallback;
	}
	const n = toNumber(arg as Scalar, ctx.date1904);
	return typeof n === 'number' ? Math.trunc(n) : n;
}

/**
 * Clamp references that run to the edge of the grid (whole columns / rows)
 * to the used range, the way the evaluator materialises them.
 */
export function clampToUsed(ref: RefValue, ctx: FunctionContext): RangeAddress {
	const r = ref.range;
	if (r.bottom !== MAX_ROW_INDEX && r.right !== MAX_COL_INDEX) {
		return r;
	}
	const used = ctx.usedRange(ref.sheet);
	const bottom =
		r.bottom === MAX_ROW_INDEX
			? Math.max(r.top, Math.min(r.bottom, used?.bottom ?? r.top))
			: r.bottom;
	const right =
		r.right === MAX_COL_INDEX
			? Math.max(r.left, Math.min(r.right, used?.right ?? r.left))
			: r.right;
	return { top: r.top, left: r.left, bottom, right };
}

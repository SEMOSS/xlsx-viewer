/**
 * Argument handling for the engineering functions. Like Excel, they reject
 * logical values (`BIN2DEC(TRUE)` is #VALUE!) instead of treating them as 1 / 0.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext } from '../../types';
import { isOmitted, numArg, type NumResult } from '../datetime/args';

/** Numeric argument; logical values are #VALUE!, blank / omitted is 0. */
export function engNum(arg: FunctionArg | undefined, ctx: FunctionContext): NumResult {
	return typeof arg === 'boolean' ? ERRORS.VALUE : numArg(arg, ctx);
}

/** Optional numeric argument with a default. */
export function optEngNum(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
	dflt: number,
): NumResult {
	return isOmitted(arg) ? dflt : engNum(arg, ctx);
}

/** Read several numeric arguments; `null` marks a required one, numbers are defaults. */
export function engNumbers(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	defaults: ReadonlyArray<number | null>,
): number[] | CellError {
	const out: number[] = [];
	for (let i = 0; i < defaults.length; i++) {
		const d = defaults[i];
		const v = d === null ? engNum(args[i], ctx) : optEngNum(args[i], ctx, d);
		if (isCellError(v)) {
			return v;
		}
		out.push(v);
	}
	return out;
}

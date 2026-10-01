/**
 * Argument coercion helpers shared by the date/time, financial and
 * engineering function modules.
 *
 * Scalar parameters arrive already dereferenced (see `invokeFunction`).
 * An argument is *omitted* when it was not supplied or written empty
 * (`PMT(r,n,pv,,1)`); empty arguments reach implementations as `null` (or a
 * `missing` AST node), the same as a blank cell. Optional arguments fall
 * back to their default in all of these cases; required numeric arguments
 * treat them as 0, as Excel does.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean, toNumber, toText } from '../../coerce';
import type { FunctionArg, FunctionContext, Scalar } from '../../types';
import { maxSerial } from './serial';

export type NumResult = number | CellError;

/** Marks a required argument in {@link readNumbers}. */
export const REQUIRED = null;

/** The argument was not supplied, or was written empty. */
export function isOmitted(arg: FunctionArg | undefined): boolean {
	if (arg === undefined || arg === null) {
		return true;
	}
	return typeof arg === 'object' && 'type' in arg && arg.type === 'missing';
}

/** Narrow an argument of a `value` parameter to a scalar (references never reach here). */
function scalarOf(arg: FunctionArg | undefined): Scalar | CellError {
	if (isOmitted(arg)) {
		return null;
	}
	if (isCellError(arg)) {
		return arg;
	}
	if (typeof arg === 'object') {
		return ERRORS.VALUE;
	}
	return arg as Scalar;
}

/** Numeric value of a scalar argument (blank or omitted is 0, numeric / date text is parsed). */
export function numArg(arg: FunctionArg | undefined, ctx: FunctionContext): NumResult {
	const v = scalarOf(arg);
	return isCellError(v) ? v : toNumber(v, ctx.date1904);
}

/** Optional numeric argument: `dflt` when omitted or blank. */
export function optNum(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
	dflt: number,
): NumResult {
	return isOmitted(arg) ? dflt : numArg(arg, ctx);
}

/**
 * Read several numeric arguments. `defaults[i]` is used when argument `i` is
 * omitted; {@link REQUIRED} (`null`) marks a required argument. Returns the
 * first coercion error instead of the list.
 */
export function readNumbers(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	defaults: ReadonlyArray<number | null>,
): number[] | CellError {
	const out: number[] = [];
	for (let i = 0; i < defaults.length; i++) {
		const dflt = defaults[i];
		const v = dflt === null ? numArg(args[i], ctx) : optNum(args[i], ctx, dflt);
		if (isCellError(v)) {
			return v;
		}
		out.push(v);
	}
	return out;
}

/** Truncated integer argument (Excel truncates toward zero). */
export function intArg(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
	dflt?: number,
): NumResult {
	const v = dflt === undefined ? numArg(arg, ctx) : optNum(arg, ctx, dflt);
	return isCellError(v) ? v : Math.trunc(v);
}

/** Optional logical argument (numbers: non-zero is TRUE; text must read TRUE / FALSE). */
export function boolArg(arg: FunctionArg | undefined, dflt: boolean): boolean | CellError {
	if (isOmitted(arg)) {
		return dflt;
	}
	const v = scalarOf(arg);
	return isCellError(v) ? v : toBoolean(v);
}

/** Text value of a scalar argument (numbers use Excel's general formatting). */
export function textArg(arg: FunctionArg | undefined): string | CellError {
	const v = scalarOf(arg);
	return isCellError(v) ? v : toText(v);
}

/**
 * A date argument as a whole serial (time of day dropped). Text is parsed as
 * a date; text that is not a date is `#VALUE!`. Serials before the start of
 * the date system or after 9999-12-31 yield `outOfRange` (`#NUM!` for the
 * date functions, `#VALUE!` for the securities functions).
 */
export function dateArg(
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
	outOfRange: CellError = ERRORS.NUM,
): NumResult {
	const n = numArg(arg, ctx);
	if (isCellError(n)) {
		return n;
	}
	const d = Math.floor(n);
	return d < 0 || d > maxSerial(ctx.date1904) ? outOfRange : d;
}

/** Read several date arguments with {@link dateArg}. */
export function readDates(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	count: number,
	outOfRange: CellError = ERRORS.NUM,
): number[] | CellError {
	const out: number[] = [];
	for (let i = 0; i < count; i++) {
		const d = dateArg(args[i], ctx, outOfRange);
		if (isCellError(d)) {
			return d;
		}
		out.push(d);
	}
	return out;
}

/** `#NUM!` for non-finite results (overflow, NaN), the value otherwise (negative zero becomes 0). */
export function finite(n: number): NumResult {
	if (!Number.isFinite(n)) {
		return ERRORS.NUM;
	}
	return n === 0 ? 0 : n;
}

import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import { toBoolean, toNumber, toText } from '../../coerce';
import type { FunctionArg, FunctionContext, Scalar } from '../../types';
import { isArrayValue } from '../../types';

/**
 * Argument helpers shared by the math, text and information function
 * modules. They normalise the two shapes an explicitly empty argument
 * (`ROUND(1.5,)`) can take: an AST `{ type: 'missing' }` node for `any` /
 * `ref` parameters and `null` for `value` parameters.
 */

/** An explicitly empty argument as delivered to `any` / `ref` parameters. */
export function isMissingArg(arg: unknown): boolean {
	if (typeof arg !== 'object' || arg === null) {
		return false;
	}
	if ((arg as { type?: unknown }).type === 'missing') {
		return true;
	}
	// A missing node that went through array materialisation.
	return (
		isArrayValue(arg) &&
		arg.rows === 1 &&
		arg.cols === 1 &&
		typeof arg.values[0][0] === 'object' &&
		arg.values[0][0] !== null &&
		(arg.values[0][0] as { type?: unknown }).type === 'missing'
	);
}

/**
 * Scalar value of argument `index`: `undefined` when the argument was not
 * supplied, `null` when it is blank or explicitly empty.
 */
export function scalarArg(args: readonly FunctionArg[], index: number): Scalar | undefined {
	if (index >= args.length) {
		return undefined;
	}
	const v = args[index];
	return isMissingArg(v) ? null : (v as Scalar);
}

/** The argument was supplied with a non-blank value (new-style optional arguments). */
export function hasValue(args: readonly FunctionArg[], index: number): boolean {
	const v = scalarArg(args, index);
	return v !== undefined && v !== null;
}

/** Numeric coercion of a `value` argument (blank / empty = 0, TRUE = 1, numeric text parsed). */
export function numberArg(v: FunctionArg | undefined, ctx: FunctionContext): number | CellError {
	if (v === undefined || isMissingArg(v)) {
		return 0;
	}
	return toNumber(v as Scalar, ctx.date1904);
}

/** Numeric argument `index`, or `fallback` when it was not supplied at all. */
export function optionalNumber(
	args: readonly FunctionArg[],
	index: number,
	ctx: FunctionContext,
	fallback: number,
): number | CellError {
	return index < args.length ? numberArg(args[index], ctx) : fallback;
}

/**
 * Numeric argument `index` of a newer function, where an explicitly empty
 * argument also selects the default (`TEXTBEFORE(A1, ",", , 1)`).
 */
export function defaultedNumber(
	args: readonly FunctionArg[],
	index: number,
	ctx: FunctionContext,
	fallback: number,
): number | CellError {
	return hasValue(args, index) ? numberArg(args[index], ctx) : fallback;
}

/** Integer argument: numeric coercion, then truncation toward zero. */
export function integerArg(v: FunctionArg | undefined, ctx: FunctionContext): number | CellError {
	const n = numberArg(v, ctx);
	return typeof n === 'number' ? Math.trunc(n) : n;
}

/** Text coercion of a `value` argument (blank / empty = ""). */
export function textArg(v: FunctionArg | undefined): string | CellError {
	if (v === undefined || isMissingArg(v)) {
		return '';
	}
	return toText(v as Scalar);
}

/** Boolean coercion of a `value` argument (blank / empty = FALSE). */
export function booleanArg(v: FunctionArg | undefined): boolean | CellError {
	if (v === undefined || isMissingArg(v)) {
		return false;
	}
	return toBoolean(v as Scalar);
}

/** A finite number, `#NUM!` otherwise; negative zero is normalised to 0. */
export function finite(r: number | CellError): number | CellError {
	if (typeof r !== 'number') {
		return r;
	}
	if (!Number.isFinite(r)) {
		return ERRORS.NUM;
	}
	return r === 0 ? 0 : r;
}

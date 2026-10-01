import type { CellError } from '../../../types/cell';
import { isCellError } from '../../../values/errors';
import { toBoolean, toNumber } from '../../coerce';
import type { ArrayValue, FunctionArg, Scalar } from '../../types';
import { isArrayValue } from '../../types';

/**
 * Argument helpers shared by the lookup, dynamic-array and database modules.
 *
 * An empty argument (`F(a,,b)`) can reach an implementation as `null` or as
 * a `missing` AST node, depending on the parameter kind; absent trailing
 * arguments are simply not in the list. These helpers treat them alike.
 */

/** The value is a `missing` AST node (an empty argument slot). */
export function isMissingNode(arg: unknown): boolean {
	return typeof arg === 'object' && arg !== null && (arg as { type?: unknown }).type === 'missing';
}

/** Argument `i` is absent, empty or blank, so an optional parameter takes its default. */
export function isOmitted(args: readonly FunctionArg[], i: number): boolean {
	const a = args[i];
	return a === undefined || a === null || isMissingNode(a);
}

/** Argument `i` of a `value` parameter as a scalar (absent or empty arguments read as blank). */
export function scalarArg(args: readonly FunctionArg[], i: number): Scalar {
	const a = args[i];
	if (a === undefined || isMissingNode(a)) {
		return null;
	}
	return a as Scalar;
}

/** Optional number: `fallback` when omitted, otherwise coerced (an error is returned as-is). */
export function optionalNumber(
	args: readonly FunctionArg[],
	i: number,
	fallback: number,
	date1904 = false,
): number | CellError {
	if (isOmitted(args, i)) {
		return fallback;
	}
	return toNumber(scalarArg(args, i), date1904);
}

/** Optional number truncated toward zero. */
export function optionalInteger(
	args: readonly FunctionArg[],
	i: number,
	fallback: number,
	date1904 = false,
): number | CellError {
	const n = optionalNumber(args, i, fallback, date1904);
	return typeof n === 'number' ? Math.trunc(n) : n;
}

/** Optional logical: `fallback` when omitted, otherwise coerced. */
export function optionalBoolean(
	args: readonly FunctionArg[],
	i: number,
	fallback: boolean,
): boolean | CellError {
	if (isOmitted(args, i)) {
		return fallback;
	}
	return toBoolean(scalarArg(args, i));
}

/**
 * An `array` parameter. The evaluator substitutes an error when the value
 * cannot be materialised (a multi-area union, a LAMBDA), so that error is
 * returned for the caller to propagate.
 */
export function arrayArg(args: readonly FunctionArg[], i: number): ArrayValue | CellError {
	const a = args[i];
	if (isCellError(a)) {
		return a;
	}
	if (isArrayValue(a)) {
		// An empty argument materialised for an `array` parameter wraps the `missing` node.
		if (a.rows === 1 && a.cols === 1 && isMissingNode(a.values[0][0])) {
			return { kind: 'array', rows: 1, cols: 1, values: [[null]] };
		}
		return a;
	}
	const v = a === undefined || isMissingNode(a) ? null : (a as Scalar);
	return { kind: 'array', rows: 1, cols: 1, values: [[v]] };
}

/** An `array` parameter that was left empty (it arrives as a 1x1 array of blank). */
export function isOmittedArray(args: readonly FunctionArg[], i: number): boolean {
	if (isOmitted(args, i)) {
		return true;
	}
	const a = args[i];
	if (isArrayValue(a) && a.rows === 1 && a.cols === 1) {
		const v: unknown = a.values[0][0];
		return v === null || v === undefined || isMissingNode(v);
	}
	return false;
}

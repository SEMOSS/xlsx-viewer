import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import {
	checkNumber,
	collectNumbers,
	type CollectNumbersOptions,
	toBoolean,
	toNumber,
} from '../../coerce';
import { toArrayValue } from '../../invoke';
import { XLFN_FUNCTIONS } from '../../prefixes';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	Scalar,
} from '../../types';

/** A numeric result or an error value. */
export type NumResult = number | CellError;

/** Attach the `_xlfn.` file prefix to functions introduced in Excel 2010 or later. */
export function define(def: FunctionDefinition): FunctionDefinition {
	return XLFN_FUNCTIONS.has(def.name) ? { ...def, filePrefix: '_xlfn.' } : def;
}

/** An omitted argument, such as the middle one in `f(1,,2)` or one not supplied at all. */
export function isMissing(arg: FunctionArg | undefined): boolean {
	return (
		arg === undefined ||
		(typeof arg === 'object' && arg !== null && 'type' in arg && arg.type === 'missing')
	);
}

/** A numeric result as a formula value: NaN and infinities become `#NUM!`. */
export function finish(r: NumResult): FormulaValue {
	return typeof r === 'number' ? checkNumber(r) : r;
}

/**
 * Scalar argument kinds: `n` number, `i` number truncated toward zero, `b`
 * logical (passed on as 1 or 0). The object forms are optional arguments
 * with the value used when they are omitted.
 */
export type ArgSpec = 'n' | 'i' | 'b' | { n: number } | { i: number } | { b: boolean };

function coerceArg(spec: ArgSpec, arg: FunctionArg | undefined, ctx: FunctionContext): NumResult {
	const omitted = isMissing(arg);
	if (typeof spec === 'object' && omitted) {
		return 'b' in spec ? (spec.b ? 1 : 0) : 'i' in spec ? spec.i : spec.n;
	}
	const kind = typeof spec === 'string' ? spec : 'b' in spec ? 'b' : 'i' in spec ? 'i' : 'n';
	const value: Scalar = omitted ? null : (arg as Scalar);
	if (kind === 'b') {
		const b = toBoolean(value);
		return typeof b === 'boolean' ? (b ? 1 : 0) : b;
	}
	const n = toNumber(value, ctx.date1904);
	if (typeof n !== 'number') {
		return n;
	}
	return kind === 'i' ? Math.trunc(n) : n;
}

/** Coerce scalar arguments by spec; the first failure is returned as an error. */
export function coerceArgs(
	spec: readonly ArgSpec[],
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	offset = 0,
): number[] | CellError {
	const out: number[] = [];
	for (let i = 0; i < spec.length; i++) {
		const v = coerceArg(spec[i], args[offset + i], ctx);
		if (typeof v !== 'number') {
			return v;
		}
		out.push(v);
	}
	return out;
}

/**
 * A function whose arguments are all scalars. Array arguments are lifted
 * element-wise by the evaluator; results that are not finite become `#NUM!`.
 */
export function scalarFunction(
	name: string,
	spec: readonly ArgSpec[],
	compute: (v: number[]) => NumResult,
): FunctionDefinition {
	const minArgs = spec.filter((s) => typeof s === 'string').length;
	return define({
		name,
		minArgs,
		maxArgs: spec.length,
		params: ['value'],
		impl: (args, ctx) => {
			const v = coerceArgs(spec, args, ctx);
			return isCellError(v) ? v : finish(compute(v));
		},
	});
}

/** Numbers of a SUM-style argument list: typed arguments coerced, references numeric-only. */
export function numbersOf(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	opts?: CollectNumbersOptions,
): number[] | CellError {
	return collectNumbers(
		args.filter((a) => !isMissing(a)),
		ctx,
		opts,
	);
}

/** A variadic aggregate over the numbers of its arguments (AVERAGE, STDEV, ...). */
export function aggregateFunction(
	name: string,
	compute: (nums: number[]) => NumResult,
	opts?: CollectNumbersOptions,
): FunctionDefinition {
	return define({
		name,
		minArgs: 1,
		maxArgs: Infinity,
		params: ['any'],
		impl: (args, ctx) => {
			const nums = numbersOf(args, ctx, opts);
			return isCellError(nums) ? nums : finish(compute(nums));
		},
	});
}

/** Materialise an argument as an array (references resolved, scalars as 1x1). */
export function arrayOf(arg: FunctionArg, ctx: FunctionContext): ArrayValue | CellError {
	if (isMissing(arg)) {
		return { kind: 'array', rows: 1, cols: 1, values: [[null]] };
	}
	if (typeof arg === 'object' && arg !== null && 'type' in arg && !isCellError(arg)) {
		return ERRORS.VALUE;
	}
	return toArrayValue(arg as FormulaValue, ctx);
}

/** Row-major values of an argument, or an error. */
export function valuesOf(arg: FunctionArg, ctx: FunctionContext): Scalar[] | CellError {
	const a = arrayOf(arg, ctx);
	if (isCellError(a)) {
		return a;
	}
	const out: Scalar[] = [];
	for (const row of a.values) {
		for (const v of row) {
			out.push(v);
		}
	}
	return out;
}

/** Numeric values of an array argument (text, logicals and blanks skipped; errors returned). */
export function arrayNumbers(arg: FunctionArg, ctx: FunctionContext): number[] | CellError {
	const values = valuesOf(arg, ctx);
	if (isCellError(values)) {
		return values;
	}
	const out: number[] = [];
	for (const v of values) {
		if (typeof v === 'number') {
			out.push(v);
		} else if (isCellError(v)) {
			return v;
		}
	}
	return out;
}

/** Vertical array of values. */
export function columnOf(values: readonly Scalar[]): ArrayValue {
	return { kind: 'array', rows: values.length, cols: 1, values: values.map((v) => [v]) };
}

import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { toScalar } from '../../invoke';
import type { ArrayValue, FunctionArg, FunctionContext, LambdaValue, Scalar } from '../../types';
import { isArrayValue, isLambdaValue } from '../../types';

/**
 * Aggregation of one group for GROUPBY / PIVOTBY: `group` holds the group's
 * values, `all` the values the group is measured against (PERCENTOF and
 * two-parameter LAMBDAs receive both).
 */
export type Aggregator = (group: readonly Scalar[], all: readonly Scalar[]) => Scalar;

/** Numbers among the values (text, logicals and blanks skipped, as in ranges); the first error otherwise. */
function numbers(values: readonly Scalar[]): number[] | CellError {
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

function total(nums: readonly number[]): number {
	let s = 0;
	for (const n of nums) {
		s += n;
	}
	return s;
}

function variance(nums: readonly number[], sample: boolean): number | CellError {
	const n = nums.length;
	if (n < (sample ? 2 : 1)) {
		return ERRORS.DIV0;
	}
	const mean = total(nums) / n;
	let ss = 0;
	for (const v of nums) {
		ss += (v - mean) * (v - mean);
	}
	return ss / (sample ? n - 1 : n);
}

/** Wrap a reducer over the group's numbers (errors in the group are returned). */
function numeric(reduce: (nums: number[]) => number | CellError): Aggregator {
	return (group) => {
		const nums = numbers(group);
		return isCellError(nums) ? nums : reduce(nums);
	};
}

function median(nums: number[]): number | CellError {
	if (nums.length === 0) {
		return ERRORS.NUM;
	}
	const sorted = nums.slice().sort((a, b) => a - b);
	const mid = sorted.length >> 1;
	return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function stdev(sample: boolean): Aggregator {
	return numeric((nums) => {
		const v = variance(nums, sample);
		return isCellError(v) ? v : Math.sqrt(v);
	});
}

/** Functions usable by name (eta-reduced) as the GROUPBY / PIVOTBY function argument. */
const NATIVE: Readonly<Record<string, Aggregator>> = {
	SUM: numeric(total),
	AVERAGE: numeric((n) => (n.length ? total(n) / n.length : ERRORS.DIV0)),
	COUNT: (g) => g.reduce<number>((c, v) => (typeof v === 'number' ? c + 1 : c), 0),
	COUNTA: (g) => g.reduce<number>((c, v) => (v !== null ? c + 1 : c), 0),
	MAX: numeric((n) => (n.length ? n.reduce((a, b) => (b > a ? b : a)) : 0)),
	MIN: numeric((n) => (n.length ? n.reduce((a, b) => (b < a ? b : a)) : 0)),
	PRODUCT: numeric((n) => (n.length ? n.reduce((a, b) => a * b, 1) : 0)),
	MEDIAN: numeric(median),
	'STDEV.S': stdev(true),
	STDEV: stdev(true),
	'STDEV.P': stdev(false),
	STDEVP: stdev(false),
	'VAR.S': numeric((n) => variance(n, true)),
	VAR: numeric((n) => variance(n, true)),
	'VAR.P': numeric((n) => variance(n, false)),
	VARP: numeric((n) => variance(n, false)),
	PERCENTOF: (group, all) => {
		const part = numbers(group);
		if (isCellError(part)) {
			return part;
		}
		const whole = numbers(all);
		if (isCellError(whole)) {
			return whole;
		}
		const denominator = total(whole);
		return denominator === 0 ? ERRORS.DIV0 : total(part) / denominator;
	},
};

function columnArray(values: readonly Scalar[]): ArrayValue {
	return { kind: 'array', rows: values.length, cols: 1, values: values.map((v) => [v]) };
}

function lambdaAggregator(fn: LambdaValue, ctx: FunctionContext): Aggregator {
	return (group, all) => {
		const args =
			fn.params.length >= 2 ? [columnArray(group), columnArray(all)] : [columnArray(group)];
		const result = ctx.callLambda(fn, args);
		if (isArrayValue(result) && (result.rows !== 1 || result.cols !== 1)) {
			return ERRORS.CALC;
		}
		return toScalar(result, ctx);
	};
}

function isNameNode(node: unknown): node is { type: 'name'; name: string } {
	return (
		typeof node === 'object' &&
		node !== null &&
		(node as { type?: unknown }).type === 'name' &&
		typeof (node as { name?: unknown }).name === 'string'
	);
}

/**
 * Resolve the function argument (a `lazy` parameter): a function name such
 * as `SUM` or `PERCENTOF` maps to a built-in aggregation unless a LET /
 * LAMBDA binding shadows it; anything else is evaluated and must yield a
 * LAMBDA. Other values are `#VALUE!`.
 */
export function resolveAggregator(arg: FunctionArg, ctx: FunctionContext): Aggregator | CellError {
	if (isNameNode(arg)) {
		const name = arg.name.toUpperCase();
		const native = NATIVE[name];
		if (native && !ctx.scope.has(name)) {
			return native;
		}
	}
	let fn: unknown = arg;
	if (typeof arg === 'object' && arg !== null && 'type' in arg && !isCellError(arg)) {
		fn = ctx.evaluate(arg as AstNode);
	}
	if (isLambdaValue(fn)) {
		return lambdaAggregator(fn, ctx);
	}
	return isCellError(fn) ? fn : ERRORS.VALUE;
}

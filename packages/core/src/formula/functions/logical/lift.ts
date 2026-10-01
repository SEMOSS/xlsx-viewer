import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { broadcastElement, implicitIntersect, type ArrayMode } from '../../invoke';
import type { ArrayValue, FormulaValue, FunctionArg, FunctionContext, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';

/**
 * Helpers for functions with `lazy` parameters (IF, IFERROR, CHOOSE, ...),
 * which evaluate their arguments themselves and therefore also perform the
 * array lifting that `invokeFunction` does for `value` parameters.
 */

/** Array mode of the evaluation (the evaluator's frames expose it; other contexts are dynamic). */
export function arrayModeOf(ctx: FunctionContext): ArrayMode {
	return (ctx as { mode?: unknown }).mode === 'legacy' ? 'legacy' : 'dynamic';
}

/** The argument is an omitted (`missing`) node or the omitted-parameter marker. */
export function isMissingNode(arg: unknown): boolean {
	return typeof arg === 'object' && arg !== null && (arg as { type?: unknown }).type === 'missing';
}

/**
 * Reduce an evaluated value to a scalar, or to an array when it holds
 * several values. In legacy mode a multi-cell reference is implicitly
 * intersected with the formula cell; arrays stay arrays (Excel evaluates
 * array constants element-wise even in formulas entered without CSE).
 */
export function liftValue(
	v: FormulaValue | FunctionArg,
	ctx: FunctionContext,
): Scalar | ArrayValue {
	if (isRefValue(v)) {
		if (v.sheetEnd !== undefined && v.sheetEnd !== v.sheet) {
			return ERRORS.VALUE;
		}
		const r = v.range;
		if (r.top === r.bottom && r.left === r.right) {
			return ctx.getCellValue(v.sheet, r.top, r.left);
		}
		if (arrayModeOf(ctx) === 'legacy') {
			return implicitIntersect(v, ctx);
		}
		return liftValue(ctx.resolveRef(v), ctx);
	}
	if (isArrayValue(v)) {
		if (v.rows === 0 || v.cols === 0) {
			return ERRORS.CALC;
		}
		return v.rows === 1 && v.cols === 1 ? v.values[0][0] : v;
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? liftValue(v.refs[0], ctx) : ERRORS.VALUE;
	}
	if (isLambdaValue(v)) {
		return ERRORS.CALC;
	}
	if (isMissingNode(v)) {
		return null;
	}
	return v as Scalar;
}

/** Evaluate a lazy argument and {@link liftValue} it. */
export function evalLifted(node: FunctionArg, ctx: FunctionContext): Scalar | ArrayValue {
	if (isMissingNode(node)) {
		return null;
	}
	return liftValue(ctx.evaluate(node as AstNode), ctx);
}

/** Value of a lazily evaluated result argument: an empty one (`IF(A1,,1)`) is 0. */
export function branchValue(node: FunctionArg, ctx: FunctionContext): FormulaValue {
	if (isMissingNode(node)) {
		return 0;
	}
	return ctx.evaluate(node as AstNode);
}

/** Element (r, c) of a lifted value (scalars repeat; arrays broadcast, `#N/A` beyond their size). */
export function elementAt(v: Scalar | ArrayValue, r: number, c: number): Scalar {
	return isArrayValue(v) ? broadcastElement(v, r, c) : v;
}

/** Broadcast shape of several lifted values. */
export function liftedShape(values: ReadonlyArray<Scalar | ArrayValue>): {
	rows: number;
	cols: number;
} {
	let rows = 1;
	let cols = 1;
	for (const v of values) {
		if (isArrayValue(v)) {
			rows = Math.max(rows, v.rows);
			cols = Math.max(cols, v.cols);
		}
	}
	return { rows, cols };
}

/** Build an array of the given shape from a per-element function. */
export function buildArray(
	rows: number,
	cols: number,
	f: (r: number, c: number) => Scalar,
): ArrayValue {
	const values: Scalar[][] = new Array(rows);
	for (let r = 0; r < rows; r++) {
		const line: Scalar[] = new Array(cols);
		for (let c = 0; c < cols; c++) {
			line[c] = f(r, c);
		}
		values[r] = line;
	}
	return { kind: 'array', rows, cols, values };
}

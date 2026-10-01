import type { CellError } from '../types/cell';
import { ERRORS, isCellError } from '../values/errors';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	ParamKind,
	RefValue,
	Scalar,
} from './types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from './types';

/** Evaluation semantics for scalar parameters given multi-cell inputs. */
export type ArrayMode = 'dynamic' | 'legacy';

/** Parameter kind for argument `index` (the last declared kind repeats). */
export function paramKindAt(def: FunctionDefinition, index: number): ParamKind {
	if (def.params.length === 0) {
		return 'value';
	}
	return def.params[Math.min(index, def.params.length - 1)];
}

/**
 * Implicit intersection of a reference with the formula cell: a single
 * column picks the cell on the same row, a single row picks the cell in the
 * same column; anything else (or no overlap) is `#VALUE!`.
 */
export function implicitIntersect(ref: RefValue, ctx: FunctionContext): Scalar {
	const r = ref.range;
	if (r.top === r.bottom && r.left === r.right) {
		return ctx.getCellValue(ref.sheet, r.top, r.left);
	}
	if (r.left === r.right && ctx.row >= r.top && ctx.row <= r.bottom) {
		return ctx.getCellValue(ref.sheet, ctx.row, r.left);
	}
	if (r.top === r.bottom && ctx.col >= r.left && ctx.col <= r.right) {
		return ctx.getCellValue(ref.sheet, r.top, ctx.col);
	}
	return ERRORS.VALUE;
}

/** Reduce any value to a scalar the way a scalar context does (legacy rules). */
export function toScalar(value: FormulaValue, ctx: FunctionContext): Scalar {
	if (isRefValue(value)) {
		return implicitIntersect(value, ctx);
	}
	if (isArrayValue(value)) {
		return value.rows > 0 && value.cols > 0 ? value.values[0][0] : ERRORS.VALUE;
	}
	if (isMultiRefValue(value)) {
		return ERRORS.VALUE;
	}
	if (isLambdaValue(value)) {
		return ERRORS.CALC;
	}
	return value;
}

/** Materialise a value as an array (references are resolved, scalars become 1x1). */
export function toArrayValue(value: FormulaValue, ctx: FunctionContext): ArrayValue | CellError {
	if (isArrayValue(value)) {
		return value;
	}
	if (isRefValue(value)) {
		return ctx.resolveRef(value);
	}
	if (isMultiRefValue(value)) {
		return value.refs.length === 1 ? ctx.resolveRef(value.refs[0]) : ERRORS.VALUE;
	}
	if (isLambdaValue(value)) {
		return ERRORS.CALC;
	}
	return { kind: 'array', rows: 1, cols: 1, values: [[value]] };
}

/** Element (r, c) of a broadcast array: single rows / columns repeat, outside the bounds is #N/A. */
export function broadcastElement(arr: ArrayValue, r: number, c: number): Scalar {
	const rr = arr.rows === 1 ? 0 : r;
	const cc = arr.cols === 1 ? 0 : c;
	if (rr >= arr.rows || cc >= arr.cols) {
		return ERRORS.NA;
	}
	return arr.values[rr][cc];
}

/** Result size when broadcasting several arrays together. */
export function broadcastShape(arrays: readonly ArrayValue[]): { rows: number; cols: number } {
	let rows = 1;
	let cols = 1;
	for (const a of arrays) {
		rows = Math.max(rows, a.rows);
		cols = Math.max(cols, a.cols);
	}
	return { rows, cols };
}

/**
 * Prepare the arguments of a function call according to its parameter kinds
 * and invoke it. This is the single source of truth for argument handling:
 * the evaluator and the function unit tests both go through it.
 *
 * `rawArgs` are evaluated arguments (references unresolved) except for
 * `lazy` parameters, which receive the AST node supplied by the caller.
 */
export function invokeFunction(
	def: FunctionDefinition,
	rawArgs: readonly FunctionArg[],
	ctx: FunctionContext,
	mode: ArrayMode = 'dynamic',
): FormulaValue {
	if (rawArgs.length < def.minArgs || rawArgs.length > def.maxArgs) {
		return ERRORS.VALUE;
	}
	const prepared: FunctionArg[] = new Array(rawArgs.length);
	const lifted: Array<{ index: number; array: ArrayValue }> = [];
	const propagate = def.propagateErrors !== false;

	for (let i = 0; i < rawArgs.length; i++) {
		const kind = paramKindAt(def, i);
		const arg = rawArgs[i];
		switch (kind) {
			case 'lazy':
			case 'any':
			case 'ref':
				prepared[i] = arg;
				break;
			case 'array': {
				const a = toArrayValue(arg as FormulaValue, ctx);
				prepared[i] = a;
				break;
			}
			case 'value': {
				const v = arg as FormulaValue;
				if (isRefValue(v)) {
					const r = v.range;
					if (r.top === r.bottom && r.left === r.right) {
						prepared[i] = ctx.getCellValue(v.sheet, r.top, r.left);
					} else if (mode === 'legacy') {
						prepared[i] = implicitIntersect(v, ctx);
					} else {
						lifted.push({ index: i, array: ctx.resolveRef(v) });
					}
				} else if (isArrayValue(v)) {
					if (v.rows === 1 && v.cols === 1) {
						prepared[i] = v.values[0][0];
					} else if (mode === 'legacy') {
						prepared[i] = toScalar(v, ctx);
					} else {
						lifted.push({ index: i, array: v });
					}
				} else if (isMultiRefValue(v)) {
					prepared[i] = ERRORS.VALUE;
				} else if (isLambdaValue(v)) {
					prepared[i] = ERRORS.CALC;
				} else {
					prepared[i] = v as Scalar;
				}
				break;
			}
		}
	}

	if (lifted.length === 0) {
		if (propagate) {
			for (let i = 0; i < prepared.length; i++) {
				if (paramKindAt(def, i) === 'value' && isCellError(prepared[i])) {
					return prepared[i] as CellError;
				}
			}
		}
		return def.impl(prepared, ctx);
	}

	// Lift: call once per element of the broadcast shape.
	const shape = broadcastShape(lifted.map((l) => l.array));
	const out: Scalar[][] = [];
	for (let r = 0; r < shape.rows; r++) {
		const row: Scalar[] = [];
		for (let c = 0; c < shape.cols; c++) {
			const args = prepared.slice();
			for (const l of lifted) {
				args[l.index] = broadcastElement(l.array, r, c);
			}
			let result: FormulaValue | undefined;
			if (propagate) {
				for (let i = 0; i < args.length; i++) {
					if (paramKindAt(def, i) === 'value' && isCellError(args[i])) {
						result = args[i] as CellError;
						break;
					}
				}
			}
			result ??= def.impl(args, ctx);
			row.push(toScalar(result, ctx));
		}
		out.push(row);
	}
	return { kind: 'array', rows: shape.rows, cols: shape.cols, values: out };
}

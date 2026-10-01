import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type {
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	LambdaValue,
} from '../../types';
import { isMissingNode } from '../logical/lift';
import { withFilePrefix } from '../logical/prefix';
import {
	acceptsArgs,
	dimensionArg,
	elementResult,
	fill,
	lambdaArg,
	MAX_DIMENSIONS,
	type Source,
	toSource,
} from './helpers';

/** MAP, REDUCE, SCAN, MAKEARRAY, BYROW and BYCOL. */

const MAP: FunctionDefinition = withFilePrefix({
	name: 'MAP',
	minArgs: 2,
	maxArgs: 254,
	params: ['any'],
	impl: (args, ctx) => {
		const fn = lambdaArg(args[args.length - 1]);
		if (isCellError(fn)) {
			return fn;
		}
		const sources: Source[] = [];
		for (const arg of args.slice(0, -1)) {
			const s = toSource(arg, ctx);
			if (isCellError(s)) {
				return s;
			}
			sources.push(s);
		}
		if (!acceptsArgs(fn, sources.length)) {
			return ERRORS.VALUE;
		}
		const rows = Math.max(...sources.map((s) => s.rows));
		const cols = Math.max(...sources.map((s) => s.cols));
		return fill(rows, cols, (r, c) =>
			elementResult(
				ctx.callLambda(
					fn,
					sources.map((s) => s.at(r, c)),
				),
				ctx,
			),
		);
	},
});

/** Split REDUCE / SCAN arguments: `([initial], array, lambda)`. */
function accumulatorArgs(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
): { initial: FormulaValue; source: Source; fn: LambdaValue | CellError } | CellError {
	const three = args.length === 3;
	const initialArg = three ? args[0] : undefined;
	const source = toSource(args[three ? 1 : 0], ctx);
	if (isCellError(source)) {
		return source;
	}
	const initial =
		initialArg === undefined || isMissingNode(initialArg) ? null : (initialArg as FormulaValue);
	return { initial, source, fn: lambdaArg(args[args.length - 1]) };
}

function accumulate(name: string, scan: boolean): FunctionDefinition {
	return withFilePrefix({
		name,
		minArgs: 2,
		maxArgs: 3,
		params: ['any'],
		impl: (args, ctx) => {
			const parts = accumulatorArgs(args, ctx);
			if (isCellError(parts)) {
				return parts;
			}
			const { source, fn } = parts;
			if (isCellError(fn)) {
				return fn;
			}
			if (!acceptsArgs(fn, 2)) {
				return ERRORS.VALUE;
			}
			let acc = parts.initial;
			if (!scan) {
				for (let r = 0; r < source.rows; r++) {
					for (let c = 0; c < source.cols; c++) {
						acc = ctx.callLambda(fn, [acc, source.at(r, c)]);
					}
				}
				return acc;
			}
			return fill(source.rows, source.cols, (r, c) => {
				acc = ctx.callLambda(fn, [acc, source.at(r, c)]);
				return elementResult(acc, ctx);
			});
		},
	});
}

const MAKEARRAY: FunctionDefinition = withFilePrefix({
	name: 'MAKEARRAY',
	minArgs: 3,
	maxArgs: 3,
	params: ['value', 'value', 'any'],
	impl: (args, ctx) => {
		const rows = dimensionArg(args[0], ctx, MAX_DIMENSIONS.rows);
		if (isCellError(rows)) {
			return rows;
		}
		const cols = dimensionArg(args[1], ctx, MAX_DIMENSIONS.cols);
		if (isCellError(cols)) {
			return cols;
		}
		const fn = lambdaArg(args[2]);
		if (isCellError(fn)) {
			return fn;
		}
		if (!acceptsArgs(fn, 2)) {
			return ERRORS.VALUE;
		}
		return fill(rows, cols, (r, c) => elementResult(ctx.callLambda(fn, [r + 1, c + 1]), ctx));
	},
});

function byLine(name: string, byRow: boolean): FunctionDefinition {
	return withFilePrefix({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['any', 'any'],
		impl: (args, ctx) => {
			const source = toSource(args[0], ctx);
			if (isCellError(source)) {
				return source;
			}
			const fn = lambdaArg(args[1]);
			if (isCellError(fn)) {
				return fn;
			}
			if (!acceptsArgs(fn, 1)) {
				return ERRORS.VALUE;
			}
			const result = (i: number) =>
				elementResult(ctx.callLambda(fn, [byRow ? source.row(i) : source.column(i)]), ctx);
			return byRow
				? fill(source.rows, 1, (r) => result(r))
				: fill(1, source.cols, (_, c) => result(c));
		},
	});
}

export const ITERATE_FUNCTIONS: FunctionDefinition[] = [
	MAP,
	accumulate('REDUCE', false),
	accumulate('SCAN', true),
	MAKEARRAY,
	byLine('BYROW', true),
	byLine('BYCOL', false),
];

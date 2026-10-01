import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext } from '../../types';
import { type Grid, gridAt, toGrid } from './counting/grid';

/** Paired numeric observations. */
export interface Pairs {
	xs: number[];
	ys: number[];
}

/** Last flat (row-major) index that can hold a value, plus one. */
function flatLimit(g: Grid): number {
	return g.filledCols === 0 ? 0 : Math.min(g.rows, g.filledRows) * g.cols;
}

function flatAt(g: Grid, i: number) {
	return gridAt(g, Math.floor(i / g.cols), i % g.cols);
}

/**
 * Pair two arrays element by element in row-major order. Both must hold
 * the same number of cells (`#N/A` otherwise); a pair is used only when
 * both values are numbers, and the first error found is returned.
 */
export function numericPairs(
	first: FunctionArg,
	second: FunctionArg,
	ctx: FunctionContext,
): Pairs | CellError {
	const a = toGrid(first, ctx);
	if (isCellError(a)) {
		return a;
	}
	const b = toGrid(second, ctx);
	if (isCellError(b)) {
		return b;
	}
	if (a.rows * a.cols !== b.rows * b.cols) {
		return ERRORS.NA;
	}
	const limit = Math.min(flatLimit(a), flatLimit(b));
	const xs: number[] = [];
	const ys: number[] = [];
	for (let i = 0; i < limit; i++) {
		const x = flatAt(a, i);
		const y = flatAt(b, i);
		if (isCellError(x)) {
			return x;
		}
		if (isCellError(y)) {
			return y;
		}
		if (typeof x === 'number' && typeof y === 'number') {
			xs.push(x);
			ys.push(y);
		}
	}
	// Errors beyond the shorter filled area still propagate.
	for (const g of [a, b]) {
		for (let i = limit; i < flatLimit(g); i++) {
			const v = flatAt(g, i);
			if (isCellError(v)) {
				return v;
			}
		}
	}
	return { xs, ys };
}

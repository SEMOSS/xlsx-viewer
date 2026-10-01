import type { CellError } from '../../../../types/cell';
import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionArg, FunctionContext } from '../../../types';
import { isMissing } from '../args';
import { gridAt, toGrid } from '../counting/grid';

/** A rectangular block of numbers. */
export interface Matrix {
	rows: number;
	cols: number;
	data: number[][];
}

/**
 * All values of an argument as numbers. Blank cells, text and logical
 * values are #VALUE! (the regression functions need complete data); the
 * first error value found is returned as is.
 */
export function numericMatrix(arg: FunctionArg, ctx: FunctionContext): Matrix | CellError {
	const g = toGrid(arg, ctx);
	if (isCellError(g)) {
		return g;
	}
	const data: number[][] = [];
	let invalid = g.filledRows < g.rows || g.filledCols < g.cols;
	for (let r = 0; r < Math.min(g.rows, g.filledRows); r++) {
		const row: number[] = [];
		for (let c = 0; c < g.cols; c++) {
			const v = gridAt(g, r, c);
			if (isCellError(v)) {
				return v;
			}
			if (typeof v !== 'number') {
				invalid = true;
			}
			row.push(typeof v === 'number' ? v : NaN);
		}
		data.push(row);
	}
	return invalid ? ERRORS.VALUE : { rows: g.rows, cols: g.cols, data };
}

/** How the x values are laid out relative to y. */
export type Layout = 'single' | 'columns' | 'rows';

export interface Design {
	/** y values in row-major order. */
	y: number[];
	/** One row per observation, one column per x variable. */
	x: number[][];
	layout: Layout;
	/** Shape of known_y's. */
	yRows: number;
	yCols: number;
	/** Number of x variables. */
	k: number;
}

function flatten(m: Matrix): number[] {
	return m.data.flat();
}

/**
 * Pair known_y's with known_x's: the same shape means one variable; a y
 * column takes each x column as a variable, a y row each x row. Any other
 * combination is #REF!. Omitted x values are 1, 2, 3, ...
 */
export function buildDesign(
	yArg: FunctionArg,
	xArg: FunctionArg | undefined,
	ctx: FunctionContext,
): Design | CellError {
	const ym = numericMatrix(yArg, ctx);
	if (isCellError(ym)) {
		return ym;
	}
	const y = flatten(ym);
	const base = { y, yRows: ym.rows, yCols: ym.cols };
	if (xArg === undefined || isMissing(xArg)) {
		return { ...base, x: y.map((_, i) => [i + 1]), layout: 'single', k: 1 };
	}
	const xm = numericMatrix(xArg, ctx);
	if (isCellError(xm)) {
		return xm;
	}
	if (xm.rows === ym.rows && xm.cols === ym.cols) {
		return { ...base, x: flatten(xm).map((v) => [v]), layout: 'single', k: 1 };
	}
	if (ym.cols === 1 && xm.rows === ym.rows) {
		return { ...base, x: xm.data.map((row) => row.slice()), layout: 'columns', k: xm.cols };
	}
	if (ym.rows === 1 && xm.cols === ym.cols) {
		const x = y.map((_, i) => xm.data.map((row) => row[i]));
		return { ...base, x, layout: 'rows', k: xm.rows };
	}
	return ERRORS.REF;
}

/** Observations of new_x's for a design, with the shape of the result. */
export interface NewPoints {
	x: number[][];
	rows: number;
	cols: number;
}

export function newPoints(
	design: Design,
	arg: FunctionArg | undefined,
	ctx: FunctionContext,
): NewPoints | CellError {
	if (arg === undefined || isMissing(arg)) {
		const shape = design.layout === 'single' ? { rows: design.yRows, cols: design.yCols } : null;
		const count = design.x.length;
		return {
			x: design.x,
			rows: shape ? shape.rows : design.layout === 'columns' ? count : 1,
			cols: shape ? shape.cols : design.layout === 'columns' ? 1 : count,
		};
	}
	const m = numericMatrix(arg, ctx);
	if (isCellError(m)) {
		return m;
	}
	if (design.layout === 'single') {
		return { x: flatten(m).map((v) => [v]), rows: m.rows, cols: m.cols };
	}
	if (design.layout === 'columns') {
		return m.cols === design.k ? { x: m.data, rows: m.rows, cols: 1 } : ERRORS.REF;
	}
	if (m.rows !== design.k) {
		return ERRORS.REF;
	}
	return {
		x: Array.from({ length: m.cols }, (_, i) => m.data.map((row) => row[i])),
		rows: 1,
		cols: m.cols,
	};
}

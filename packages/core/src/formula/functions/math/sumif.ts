import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../../address/constants';
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { createCriteriaMatcher } from '../../criteria';
import type {
	ArrayValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	ParamKind,
	RefValue,
	Scalar,
} from '../../types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingArg, scalarArg } from './args';
import { CompensatedSum } from './sums';

/**
 * A criteria / sum range: its declared size plus the materialised values,
 * which may be smaller when a whole-column reference was clamped to the
 * used range (cells outside the data are blank).
 */
interface Grid {
	rows: number;
	cols: number;
	data: ArrayValue;
}

function refGrid(ref: RefValue, ctx: FunctionContext): Grid | CellError {
	if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
		return ERRORS.VALUE;
	}
	const r = ref.range;
	return { rows: r.bottom - r.top + 1, cols: r.right - r.left + 1, data: ctx.resolveRef(ref) };
}

/** Grid of a range argument; references and arrays are accepted, anything else is #VALUE!. */
function toGrid(arg: FunctionArg, ctx: FunctionContext): Grid | CellError {
	if (isRefValue(arg)) {
		return refGrid(arg, ctx);
	}
	if (isMultiRefValue(arg)) {
		return arg.refs.length === 1 ? refGrid(arg.refs[0], ctx) : ERRORS.VALUE;
	}
	if (isArrayValue(arg)) {
		return { rows: arg.rows, cols: arg.cols, data: arg };
	}
	return isCellError(arg) ? arg : ERRORS.VALUE;
}

/** SUMIF's sum range: its top-left cell, resized to the shape of the criteria range. */
function resizedGrid(
	arg: FunctionArg,
	rows: number,
	cols: number,
	ctx: FunctionContext,
): Grid | CellError {
	if (isRefValue(arg)) {
		const { top, left } = arg.range;
		const range = {
			top,
			left,
			bottom: Math.min(MAX_ROW_INDEX, top + rows - 1),
			right: Math.min(MAX_COL_INDEX, left + cols - 1),
		};
		const grid = refGrid({ ...arg, range }, ctx);
		return isCellError(grid) ? grid : { ...grid, rows, cols };
	}
	return toGrid(arg, ctx);
}

function cellAt(grid: Grid, r: number, c: number): Scalar {
	return r < grid.data.rows && c < grid.data.cols ? grid.data.values[r][c] : null;
}

/**
 * Rows / columns worth visiting within the criteria shape: beyond every
 * grid's materialised data all cells are blank and contribute nothing.
 */
function extent(grids: readonly Grid[], shape: Grid): { rows: number; cols: number } {
	let rows = 0;
	let cols = 0;
	for (const g of grids) {
		rows = Math.max(rows, g.data.rows);
		cols = Math.max(cols, g.data.cols);
	}
	return { rows: Math.min(rows, shape.rows), cols: Math.min(cols, shape.cols) };
}

interface Condition {
	grid: Grid;
	match: (v: Scalar) => boolean;
}

/** Sum the numbers of `sum` where every condition matches; errors in summed cells propagate. */
function conditionalSum(sum: Grid, conditions: readonly Condition[]): number | CellError {
	const { rows, cols } = extent([sum, ...conditions.map((c) => c.grid)], conditions[0].grid);
	const acc = new CompensatedSum();
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			if (!conditions.every((cond) => cond.match(cellAt(cond.grid, r, c)))) {
				continue;
			}
			const v = cellAt(sum, r, c);
			if (typeof v === 'number') {
				acc.add(v);
			} else if (isCellError(v)) {
				return v;
			}
		}
	}
	const total = acc.value;
	return Number.isFinite(total) ? total : ERRORS.NUM;
}

const SUMIF: FunctionDefinition = {
	name: 'SUMIF',
	minArgs: 2,
	maxArgs: 3,
	params: ['ref', 'value', 'ref'],
	propagateErrors: false,
	impl: (args, ctx) => {
		const range = toGrid(args[0], ctx);
		if (isCellError(range)) {
			return range;
		}
		let sum = range;
		if (args.length > 2 && args[2] !== null && !isMissingArg(args[2])) {
			const resized = resizedGrid(args[2], range.rows, range.cols, ctx);
			if (isCellError(resized)) {
				return resized;
			}
			sum = resized;
		}
		const match = createCriteriaMatcher(scalarArg(args, 1) ?? null, ctx.date1904);
		return conditionalSum(sum, [{ grid: range, match }]);
	},
};

/** SUMIFS(sum_range, criteria_range1, criteria1, ...): ranges then criteria, alternating. */
const SUMIFS_PARAMS: ParamKind[] = ['ref'];
for (let i = 0; i < 127; i++) {
	SUMIFS_PARAMS.push('ref', 'value');
}

const SUMIFS: FunctionDefinition = {
	name: 'SUMIFS',
	minArgs: 3,
	maxArgs: 255,
	params: SUMIFS_PARAMS,
	propagateErrors: false,
	impl: (args, ctx) => {
		if (args.length % 2 === 0) {
			return ERRORS.VALUE;
		}
		const sum = toGrid(args[0], ctx);
		if (isCellError(sum)) {
			return sum;
		}
		const conditions: Condition[] = [];
		for (let i = 1; i < args.length; i += 2) {
			const grid = toGrid(args[i], ctx);
			if (isCellError(grid)) {
				return grid;
			}
			if (grid.rows !== sum.rows || grid.cols !== sum.cols) {
				return ERRORS.VALUE;
			}
			conditions.push({
				grid,
				match: createCriteriaMatcher(scalarArg(args, i + 1) ?? null, ctx.date1904),
			});
		}
		return conditionalSum(sum, conditions);
	},
};

export const SUMIF_FUNCTIONS: FunctionDefinition[] = [SUMIF, SUMIFS];

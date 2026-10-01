import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../../address/constants';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FormulaValue, FunctionDefinition } from '../../types';
import { scalarArg } from './args';
import type { Grid } from './grid';
import { toGrid } from './grid';
import { approximateMatch } from './search';

/**
 * Element `pos` of LOOKUP's result_vector. Like Excel, a reference that is
 * shorter than the lookup vector is extended past its end (so a single cell
 * works as the start of a result column or row); a single cell extends in
 * the lookup vector's direction. A 2-D result_vector is `#N/A`.
 */
function resultAt(result: Grid, pos: number, vertical: boolean): FormulaValue {
	if (result.rows > 1 && result.cols > 1) {
		return ERRORS.NA;
	}
	const down = result.rows > 1 || (result.cols === 1 && vertical);
	const size = down ? result.rows : result.cols;
	if (pos >= size) {
		if (!result.ref) {
			return ERRORS.NA;
		}
		const { top, left } = result.ref.range;
		if (down ? top + pos > MAX_ROW_INDEX : left + pos > MAX_COL_INDEX) {
			return ERRORS.NA;
		}
	}
	return down ? result.get(pos, 0) : result.get(0, pos);
}

/**
 * LOOKUP(lookup_value, lookup_vector, [result_vector]) and
 * LOOKUP(lookup_value, array)
 *
 * Always an approximate match (Excel's binary search for the largest value
 * <= lookup_value; values of other types, blanks and errors are skipped).
 * The array form searches the first row when the array is wider than tall,
 * otherwise the first column, and returns from the last row / column. The
 * vector form searches lookup_vector (a 2-D one is read like the array form)
 * and returns the element at the same position of result_vector.
 */
export const LOOKUP: FunctionDefinition = {
	name: 'LOOKUP',
	minArgs: 2,
	maxArgs: 3,
	params: ['value', 'any', 'any'],
	impl: (args, ctx) => {
		const x = scalarArg(args, 0);
		const grid = toGrid(args[1], ctx);
		if (isCellError(grid)) {
			return grid;
		}
		const byRow = grid.cols > grid.rows;
		const keys = byRow ? grid.row(0) : grid.column(0);
		if (args.length < 3) {
			const pos = approximateMatch(keys, x);
			if (pos < 0) {
				return ERRORS.NA;
			}
			return byRow ? grid.get(grid.rows - 1, pos) : grid.get(pos, grid.cols - 1);
		}
		const result = toGrid(args[2], ctx);
		if (isCellError(result)) {
			return result;
		}
		const pos = approximateMatch(keys, x);
		return pos < 0 ? ERRORS.NA : resultAt(result, pos, !byRow);
	},
};

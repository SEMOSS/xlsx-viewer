import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean, toNumber } from '../../coerce';
import type { FormulaValue, FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { scalarArg } from './args';
import { exactMatcher } from './compare';
import { toGrid } from './grid';
import { exactSearch } from './exact-index';
import { approximateMatch } from './search';

/** Longest lookup text the legacy lookup functions accept (`#VALUE!` beyond). */
export const MAX_LOOKUP_TEXT = 255;

/**
 * Shared body of VLOOKUP (search the first column, return from a column)
 * and HLOOKUP (search the first row, return from a row).
 *
 * - The index is truncated; below 1 is `#VALUE!`, past the table is `#REF!`.
 * - range_lookup defaults to TRUE when absent; an empty argument
 *   (`VLOOKUP(x, t, 2, )`) is FALSE, as in Excel.
 * - Approximate match is Excel's binary search (sorted data assumed);
 *   exact match scans in order and supports wildcards for text.
 * - The table stays a lazy reference, so only the key column (or row) and
 *   the single result cell are read.
 */
function tableLookup(args: FunctionArg[], ctx: FunctionContext, vertical: boolean): FormulaValue {
	const x = scalarArg(args, 0);
	const grid = toGrid(args[1], ctx);
	if (isCellError(grid)) {
		return grid;
	}
	const rawIndex = toNumber(scalarArg(args, 2), ctx.date1904);
	if (isCellError(rawIndex)) {
		return rawIndex;
	}
	const index = Math.trunc(rawIndex);
	if (index < 1) {
		return ERRORS.VALUE;
	}
	if (index > (vertical ? grid.cols : grid.rows)) {
		return ERRORS.REF;
	}
	let approximate = true;
	if (args.length > 3) {
		const b = toBoolean(scalarArg(args, 3));
		if (isCellError(b)) {
			return b;
		}
		approximate = b;
	}
	if (typeof x === 'string' && x.length > MAX_LOOKUP_TEXT) {
		return ERRORS.VALUE;
	}
	const keys = vertical ? grid.column(0) : grid.row(0);
	const pos = approximate ? approximateMatch(keys, x) : exactSearch(keys, x, { wildcards: true });
	if (pos < 0) {
		return ERRORS.NA;
	}
	return vertical ? grid.get(pos, index - 1) : grid.get(index - 1, pos);
}

export const VLOOKUP: FunctionDefinition = {
	name: 'VLOOKUP',
	minArgs: 3,
	maxArgs: 4,
	params: ['value', 'any', 'value', 'value'],
	impl: (args, ctx) => tableLookup(args, ctx, true),
};

export const HLOOKUP: FunctionDefinition = {
	name: 'HLOOKUP',
	minArgs: 3,
	maxArgs: 4,
	params: ['value', 'any', 'value', 'value'],
	impl: (args, ctx) => tableLookup(args, ctx, false),
};

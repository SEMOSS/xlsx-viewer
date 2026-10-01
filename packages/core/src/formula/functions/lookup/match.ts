import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type { FunctionDefinition } from '../../types';
import { scalarArg } from './args';
import { exactMatcher } from './compare';
import { toGrid } from './grid';
import { exactSearch } from './exact-index';
import { approximateMatch } from './search';
import { MAX_LOOKUP_TEXT } from './vlookup';

/**
 * MATCH(lookup_value, lookup_array, [match_type])
 *
 * - match_type 1 (default when absent): largest value <= lookup_value, by
 *   binary search over ascending data.
 * - 0 (also an empty argument): first exact match; text uses wildcards.
 * - -1: smallest value >= lookup_value, by binary search over descending data.
 *
 * Any positive / negative match_type behaves as 1 / -1. A lookup_array that
 * is neither a single row nor a single column is `#N/A`.
 */
export const MATCH: FunctionDefinition = {
	name: 'MATCH',
	minArgs: 2,
	maxArgs: 3,
	params: ['value', 'any', 'value'],
	impl: (args, ctx) => {
		const x = scalarArg(args, 0);
		const grid = toGrid(args[1], ctx);
		if (isCellError(grid)) {
			return grid;
		}
		let type = 1;
		if (args.length > 2) {
			const n = toNumber(scalarArg(args, 2), ctx.date1904);
			if (isCellError(n)) {
				return n;
			}
			type = n > 0 ? 1 : n < 0 ? -1 : 0;
		}
		if (grid.rows > 1 && grid.cols > 1) {
			return ERRORS.NA;
		}
		if (typeof x === 'string' && x.length > MAX_LOOKUP_TEXT) {
			return ERRORS.VALUE;
		}
		const vec = grid.cols === 1 ? grid.column(0) : grid.row(0);
		const pos =
			type === 0 ? exactSearch(vec, x, { wildcards: true }) : approximateMatch(vec, x, type < 0);
		return pos < 0 ? ERRORS.NA : pos + 1;
	},
};

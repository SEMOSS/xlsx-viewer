import { ERRORS, isCellError } from '../../../values/errors';
import type { FormulaValue, FunctionDefinition } from '../../types';
import { isOmitted, scalarArg } from './args';
import {
	gridColumnResult,
	gridRowResult,
	isFullHeight,
	isFullWidth,
	isVectorGrid,
	sameExtent,
	toGrid,
} from './grid';
import { readXModes, xsearch } from './xsearch';

/**
 * XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
 *
 * lookup_array is a single column or row. A vertical lookup returns the
 * matching row of return_array (all its columns), a horizontal one the
 * matching column, so a 2-D return_array yields a row / column that spills.
 * The extents must agree, otherwise `#VALUE!`. When return_array is a
 * reference the result is a reference too, so `XLOOKUP(...):XLOOKUP(...)`
 * and `SUM(XLOOKUP(...))` work as in Excel. if_not_found replaces `#N/A`
 * only when nothing matches.
 */
export const XLOOKUP: FunctionDefinition = {
	name: 'XLOOKUP',
	minArgs: 3,
	maxArgs: 6,
	params: ['value', 'any', 'any', 'any', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const x = scalarArg(args, 0);
		const lookup = toGrid(args[1], ctx);
		if (isCellError(lookup)) {
			return lookup;
		}
		const ret = toGrid(args[2], ctx);
		if (isCellError(ret)) {
			return ret;
		}
		const modes = readXModes(args, 4, ctx.date1904);
		if (isCellError(modes)) {
			return modes;
		}
		if (!isVectorGrid(lookup)) {
			return ERRORS.VALUE;
		}
		let vertical: boolean;
		if (lookup.rows !== lookup.cols) {
			vertical = lookup.cols === 1;
		} else {
			// A single-cell lookup_array pairs with a single-row or single-column return_array.
			vertical = ret.rows === 1 || ret.cols !== 1;
		}
		const fits = vertical
			? sameExtent(lookup, lookup.rows, ret, ret.rows, isFullHeight)
			: sameExtent(lookup, lookup.cols, ret, ret.cols, isFullWidth);
		if (!fits) {
			return ERRORS.VALUE;
		}
		const pos = xsearch(vertical ? lookup.column(0) : lookup.row(0), x, modes);
		if (isCellError(pos)) {
			return pos;
		}
		if (pos < 0) {
			return isOmitted(args, 3) ? ERRORS.NA : (args[3] as FormulaValue);
		}
		return vertical ? gridRowResult(ret, pos) : gridColumnResult(ret, pos);
	},
};

/**
 * XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
 *
 * 1-based position of the match in a single row or column, with XLOOKUP's
 * match and search modes. A 2-D lookup_array is `#VALUE!`.
 */
export const XMATCH: FunctionDefinition = {
	name: 'XMATCH',
	minArgs: 2,
	maxArgs: 4,
	params: ['value', 'any', 'value', 'value'],
	filePrefix: '_xlfn.',
	impl: (args, ctx) => {
		const x = scalarArg(args, 0);
		const lookup = toGrid(args[1], ctx);
		if (isCellError(lookup)) {
			return lookup;
		}
		const modes = readXModes(args, 2, ctx.date1904);
		if (isCellError(modes)) {
			return modes;
		}
		if (!isVectorGrid(lookup)) {
			return ERRORS.VALUE;
		}
		const vec = lookup.cols === 1 ? lookup.column(0) : lookup.row(0);
		const pos = xsearch(vec, x, modes);
		if (isCellError(pos)) {
			return pos;
		}
		return pos < 0 ? ERRORS.NA : pos + 1;
	},
};

import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean } from '../../coerce';
import type { FormulaValue, FunctionDefinition } from '../../types';
import { arrayArg, isOmitted } from '../lookup/args';
import { pickCols, pickRows } from './shared';

/**
 * FILTER(array, include, [if_empty])
 *
 * include is a single column as tall as array (keeps rows) or a single row
 * as wide as array (keeps columns); any other shape is `#VALUE!`. Its values
 * are read as logicals; an error is returned as-is and text that is not
 * TRUE / FALSE is `#VALUE!`. When nothing is kept the result is if_empty,
 * or `#CALC!` (Excel's empty-array error) when it is omitted.
 */
export const FILTER: FunctionDefinition = {
	name: 'FILTER',
	minArgs: 2,
	maxArgs: 3,
	params: ['array', 'array', 'any'],
	filePrefix: '_xlfn._xlws.',
	returnsArray: true,
	impl: (args) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const include = arrayArg(args, 1);
		if (isCellError(include)) {
			return include;
		}
		let byRows: boolean;
		if (include.cols === 1 && include.rows === a.rows) {
			byRows = true;
		} else if (include.rows === 1 && include.cols === a.cols) {
			byRows = false;
		} else {
			return ERRORS.VALUE;
		}
		const n = byRows ? a.rows : a.cols;
		const keep: number[] = [];
		for (let i = 0; i < n; i++) {
			const flag = toBoolean(byRows ? include.values[i][0] : include.values[0][i]);
			if (isCellError(flag)) {
				return flag;
			}
			if (flag) {
				keep.push(i);
			}
		}
		if (keep.length === 0) {
			return isOmitted(args, 2) ? ERRORS.CALC : (args[2] as FormulaValue);
		}
		return byRows ? pickRows(a, keep) : pickCols(a, keep);
	},
};

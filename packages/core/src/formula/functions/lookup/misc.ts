import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition, Scalar } from '../../types';
import { arrayArg, isMissingNode, scalarArg } from './args';

/** Longest link_location HYPERLINK accepts. */
const MAX_LINK_LENGTH = 255;

/** TRANSPOSE(array): rows become columns. */
export const TRANSPOSE: FunctionDefinition = {
	name: 'TRANSPOSE',
	minArgs: 1,
	maxArgs: 1,
	params: ['array'],
	returnsArray: true,
	impl: (args) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const values: Scalar[][] = new Array(a.cols);
		for (let c = 0; c < a.cols; c++) {
			const line: Scalar[] = new Array(a.rows);
			for (let r = 0; r < a.rows; r++) {
				line[r] = a.values[r][c];
			}
			values[c] = line;
		}
		return { kind: 'array', rows: a.cols, cols: a.rows, values };
	},
};

/**
 * HYPERLINK(link_location, [friendly_name]): the cell shows friendly_name
 * when given, otherwise the link text. Following the link is a UI concern.
 */
export const HYPERLINK: FunctionDefinition = {
	name: 'HYPERLINK',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args) => {
		const link = scalarArg(args, 0);
		if (typeof link === 'string' && link.length > MAX_LINK_LENGTH) {
			return ERRORS.VALUE;
		}
		if (args.length < 2 || isMissingNode(args[1])) {
			return link;
		}
		return scalarArg(args, 1);
	},
};

/** GETPIVOTDATA(data_field, pivot_table, [field, item], ...): pivot tables are not evaluated, so `#REF!`. */
export const GETPIVOTDATA: FunctionDefinition = {
	name: 'GETPIVOTDATA',
	minArgs: 2,
	maxArgs: 254,
	params: ['value', 'any', 'value'],
	impl: () => ERRORS.REF,
};

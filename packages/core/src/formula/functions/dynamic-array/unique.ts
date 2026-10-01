import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition, Scalar } from '../../types';
import { arrayArg, optionalBoolean } from '../lookup/args';
import { pickCols, pickRows } from './shared';

/** Identity of a value for UNIQUE and grouping: case-insensitive text, no coercion between types. */
export function valueKey(v: Scalar): string {
	switch (typeof v) {
		case 'number':
			return `n${v}`;
		case 'string':
			return `s${v.toLowerCase()}`;
		case 'boolean':
			return v ? 'T' : 'F';
		default:
			return v === null ? 'z' : `e${v.code}`;
	}
}

/** Identity of a whole row / column (length-prefixed parts, so keys cannot collide). */
function lineKey(values: readonly Scalar[]): string {
	if (values.length === 1) {
		return valueKey(values[0]);
	}
	let key = '';
	for (const v of values) {
		const k = valueKey(v);
		key += `${k.length}:${k}`;
	}
	return key;
}

/**
 * UNIQUE(array, [by_col], [exactly_once])
 *
 * Distinct rows (or columns with by_col TRUE) in order of first appearance,
 * comparing text case-insensitively like Excel (the first spelling wins).
 * With exactly_once TRUE only rows that occur once are returned. An empty
 * result is `#CALC!`.
 */
export const UNIQUE: FunctionDefinition = {
	name: 'UNIQUE',
	minArgs: 1,
	maxArgs: 3,
	params: ['array', 'value', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args) => {
		const a = arrayArg(args, 0);
		if (isCellError(a)) {
			return a;
		}
		const byCol = optionalBoolean(args, 1, false);
		if (isCellError(byCol)) {
			return byCol;
		}
		const once = optionalBoolean(args, 2, false);
		if (isCellError(once)) {
			return once;
		}
		const count = byCol ? a.cols : a.rows;
		const groups = new Map<string, { index: number; count: number }>();
		const order: Array<{ index: number; count: number }> = [];
		const line: Scalar[] = new Array(byCol ? a.rows : a.cols);
		for (let i = 0; i < count; i++) {
			for (let j = 0; j < line.length; j++) {
				line[j] = byCol ? a.values[j][i] : a.values[i][j];
			}
			const key = lineKey(line);
			const group = groups.get(key);
			if (group) {
				group.count++;
			} else {
				const fresh = { index: i, count: 1 };
				groups.set(key, fresh);
				order.push(fresh);
			}
		}
		const keep: number[] = [];
		for (const g of order) {
			if (!once || g.count === 1) {
				keep.push(g.index);
			}
		}
		if (keep.length === 0) {
			return ERRORS.CALC;
		}
		return byCol ? pickCols(a, keep) : pickRows(a, keep);
	},
};

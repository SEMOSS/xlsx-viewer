import { compareText } from '../../coerce';
import type { Scalar } from '../../types';

/** Rank of a value in Excel's sort order; blanks get the highest rank. */
function sortRank(v: Scalar): number {
	switch (typeof v) {
		case 'number':
			return 0;
		case 'string':
			return 1;
		case 'boolean':
			return 2;
		default:
			return v === null ? 4 : 3;
	}
}

/**
 * Excel's sort order (Data > Sort, SORT, SORTBY). Ascending: numbers, then
 * text (case-insensitive), then logicals (FALSE before TRUE), then errors
 * (all equal). Descending reverses that order. Blank cells always sort
 * last, in both directions.
 */
export function compareForSort(a: Scalar, b: Scalar, descending: boolean): number {
	const ra = sortRank(a);
	const rb = sortRank(b);
	if (ra === 4 || rb === 4) {
		return ra === rb ? 0 : ra === 4 ? 1 : -1;
	}
	let c: number;
	if (ra !== rb) {
		c = ra - rb;
	} else if (ra === 0) {
		c = (a as number) < (b as number) ? -1 : (a as number) > (b as number) ? 1 : 0;
	} else if (ra === 1) {
		c = a === b ? 0 : compareText(a as string, b as string);
	} else if (ra === 2) {
		c = a === b ? 0 : a ? 1 : -1;
	} else {
		c = 0;
	}
	return descending ? -c : c;
}

/**
 * Stable permutation of `count` items ordered by several keys: `keys[k][i]`
 * is the k-th sort key of item i and `descending[k]` its direction. Ties
 * keep their original order.
 */
export function sortPermutation(
	count: number,
	keys: readonly (readonly Scalar[])[],
	descending: readonly boolean[],
): number[] {
	const perm: number[] = new Array(count);
	for (let i = 0; i < count; i++) {
		perm[i] = i;
	}
	perm.sort((x, y) => {
		for (let k = 0; k < keys.length; k++) {
			const c = compareForSort(keys[k][x], keys[k][y], descending[k]);
			if (c !== 0) {
				return c;
			}
		}
		return x - y;
	});
	return perm;
}

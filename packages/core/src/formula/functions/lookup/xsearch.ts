import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, Scalar } from '../../types';
import { optionalInteger } from './args';
import { compareSameType, exactMatcher, orderAgainst, regexMatcher } from './compare';
import type { Vec } from './grid';
import { exactSearch } from './exact-index';
import { binarySearchFirst, binarySearchLast, linearSearch } from './search';

/** XLOOKUP / XMATCH match_mode: exact, exact-or-smaller, exact-or-larger, wildcard, regex. */
export type XMatchMode = 0 | -1 | 1 | 2 | 3;
/** XLOOKUP / XMATCH search_mode: first-to-last, last-to-first, binary ascending / descending. */
export type XSearchMode = 1 | -1 | 2 | -2;

export interface XModes {
	match: XMatchMode;
	search: XSearchMode;
}

/** Read and validate match_mode / search_mode (truncated; anything else is `#VALUE!`). */
export function readXModes(
	args: readonly FunctionArg[],
	matchIndex: number,
	date1904: boolean,
): XModes | CellError {
	const match = optionalInteger(args, matchIndex, 0, date1904);
	if (isCellError(match)) {
		return match;
	}
	const search = optionalInteger(args, matchIndex + 1, 1, date1904);
	if (isCellError(search)) {
		return search;
	}
	if (match < -1 || match > 3) {
		return ERRORS.VALUE;
	}
	if (search !== 1 && search !== -1 && search !== 2 && search !== -2) {
		return ERRORS.VALUE;
	}
	return { match: match as XMatchMode, search: search as XSearchMode };
}

/**
 * Linear exact-or-next search (match_mode -1 / 1): the first exact match in
 * search order wins; otherwise the closest smaller (or larger) value of the
 * lookup value's type, ties going to the first one met in search order.
 */
function linearNearest(vec: Vec, x: Scalar, mode: -1 | 1, reverse: boolean): number {
	const exact = exactMatcher(x, { blankMatchesBlank: true });
	const order = orderAgainst(x);
	const n = vec.dataLength;
	const blankTail = n < vec.length && exact(null);
	if (reverse && blankTail) {
		return vec.length - 1;
	}
	let best = -1;
	let bestValue: Scalar = null;
	for (let k = 0; k < n; k++) {
		const i = reverse ? n - 1 - k : k;
		const v = vec.get(i);
		if (exact(v)) {
			return i;
		}
		if (!order.comparable(v)) {
			continue;
		}
		const s = order.compare(v);
		if (mode < 0 ? s >= 0 : s <= 0) {
			continue;
		}
		if (
			best < 0 ||
			(mode < 0 ? compareSameType(v, bestValue) > 0 : compareSameType(v, bestValue) < 0)
		) {
			best = i;
			bestValue = v;
		}
	}
	return blankTail ? n : best;
}

/**
 * The search behind XLOOKUP and XMATCH. Returns the 0-based index, -1 when
 * nothing matches, or `#VALUE!` for wildcard / regex matching combined with
 * a binary search (which Excel does not support).
 */
export function xsearch(vec: Vec, x: Scalar, modes: XModes): number | CellError {
	const { match, search } = modes;
	const binary = search === 2 || search === -2;
	if (match === 2 || match === 3) {
		if (binary) {
			return ERRORS.VALUE;
		}
		const m =
			match === 2 ? exactMatcher(x, { wildcards: true, blankMatchesBlank: true }) : regexMatcher(x);
		return isCellError(m) ? m : linearSearch(vec, m, search === -1);
	}
	if (!binary) {
		if (match === 0) {
			return exactSearch(vec, x, { blankMatchesBlank: true }, search === -1);
		}
		return linearNearest(vec, x, match, search === -1);
	}
	const order = orderAgainst(x);
	const ascending = search === 2;
	if (match === 0) {
		const i = binarySearchFirst(vec, order, ascending ? (s) => s >= 0 : (s) => s <= 0);
		return i >= 0 && exactMatcher(x)(vec.get(i)) ? i : -1;
	}
	if (match === -1) {
		return ascending
			? binarySearchLast(vec, order, (s) => s <= 0)
			: binarySearchFirst(vec, order, (s) => s <= 0);
	}
	return ascending
		? binarySearchFirst(vec, order, (s) => s >= 0)
		: binarySearchLast(vec, order, (s) => s >= 0);
}

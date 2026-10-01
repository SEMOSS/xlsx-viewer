import { hasWildcards } from '../../criteria';
import type { Scalar } from '../../types';
import { exactMatcher, type ExactOptions } from './compare';
import type { Vec } from './grid';
import { linearSearch } from './search';

/** First and last position of every distinct value of a vector. */
export type ExactIndex = Map<string, { first: number; last: number }>;

/** Hash key of a value under Excel's exact-match equality (text ignores case). */
function indexKey(v: Scalar): string | undefined {
	switch (typeof v) {
		case 'number':
			return `n${v}`;
		case 'boolean':
			return v ? 'bT' : 'bF';
		case 'string':
			return v === '' ? undefined : `s${v.toLowerCase()}`;
		default:
			return undefined;
	}
}

/** Build the index of a vector's data (blanks and errors are not indexed). */
export function buildExactIndex(vec: Vec): ExactIndex {
	const index: ExactIndex = new Map();
	for (let i = 0; i < vec.dataLength; i++) {
		const key = indexKey(vec.get(i));
		if (key === undefined) {
			continue;
		}
		const hit = index.get(key);
		if (hit) {
			hit.last = i;
		} else {
			index.set(key, { first: i, last: i });
		}
	}
	return index;
}

/** Below this many cells a scan beats building (or fetching) an index. */
const INDEX_MIN_LENGTH = 32;

/**
 * Exact-match search (first match, or last with `reverse`). Large reference
 * vectors use a hash index memoised by the calculation engine, so filling
 * VLOOKUP / MATCH / XLOOKUP down a table stays linear overall; wildcard and
 * blank lookups scan.
 */
export function exactSearch(vec: Vec, x: Scalar, opts: ExactOptions = {}, reverse = false): number {
	const indexable =
		vec.exactIndex !== undefined &&
		vec.dataLength >= INDEX_MIN_LENGTH &&
		(typeof x === 'number' ||
			typeof x === 'boolean' ||
			(typeof x === 'string' && x !== '' && !(opts.wildcards && hasWildcards(x))));
	if (!indexable) {
		return linearSearch(vec, exactMatcher(x, opts), reverse);
	}
	const plain = typeof x === 'string' && opts.wildcards ? x.replace(/~([*?~])/g, '$1') : x;
	const key = indexKey(plain);
	const hit = key === undefined ? undefined : vec.exactIndex!().get(key);
	if (!hit) {
		return -1;
	}
	return reverse ? hit.last : hit.first;
}

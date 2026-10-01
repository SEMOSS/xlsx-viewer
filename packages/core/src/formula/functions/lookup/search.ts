import type { Scalar } from '../../types';
import type { Matcher, Order } from './compare';
import { orderAgainst } from './compare';
import type { Vec } from './grid';

/**
 * Search primitives over a lookup vector. Every function returns a 0-based
 * index, or -1 when nothing qualifies.
 */

/**
 * Linear search in either direction. Elements past the vector's data length
 * are blank, so they are only considered when the matcher accepts a blank.
 */
export function linearSearch(vec: Vec, match: Matcher, reverse = false): number {
	const n = vec.dataLength;
	const blankTail = n < vec.length && match(null);
	if (reverse) {
		if (blankTail) {
			return vec.length - 1;
		}
		for (let i = n - 1; i >= 0; i--) {
			if (match(vec.get(i))) {
				return i;
			}
		}
		return -1;
	}
	for (let i = 0; i < n; i++) {
		if (match(vec.get(i))) {
			return i;
		}
	}
	return blankTail ? n : -1;
}

/** Probed element of a binary search step (`index` -1 when nothing is comparable). */
interface Probe {
	index: number;
	value: Scalar;
}

/**
 * Excel-style binary search probe: the first element at or after `mid`
 * (up to `hi`) whose type matches the lookup value. Blanks, errors and
 * values of other types are stepped over, which is what makes idioms such
 * as `LOOKUP(2, 1/(A:A<>""), A:A)` and `MATCH(9.99E+307, A:A)` work. Each
 * cell is read once.
 */
function probe(vec: Vec, order: Order, mid: number, hi: number, out: Probe): Probe {
	for (let j = mid; j <= hi; j++) {
		const v = vec.get(j);
		if (order.comparable(v)) {
			out.index = j;
			out.value = v;
			return out;
		}
	}
	out.index = -1;
	return out;
}

/**
 * Binary search for the last comparable element satisfying `pred`, where
 * `pred` holds on a prefix of the (assumed sorted) comparable elements.
 * `pred` receives the sign of `element - lookup value`.
 */
export function binarySearchLast(vec: Vec, order: Order, pred: (sign: number) => boolean): number {
	let lo = 0;
	let hi = vec.dataLength - 1;
	let found = -1;
	const p: Probe = { index: -1, value: null };
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1;
		const j = probe(vec, order, mid, hi, p).index;
		if (j < 0) {
			hi = mid - 1;
		} else if (pred(order.compare(p.value))) {
			found = j;
			lo = j + 1;
		} else {
			hi = mid - 1;
		}
	}
	return found;
}

/**
 * Binary search for the first comparable element satisfying `pred`, where
 * `pred` holds on a suffix of the (assumed sorted) comparable elements.
 */
export function binarySearchFirst(vec: Vec, order: Order, pred: (sign: number) => boolean): number {
	let lo = 0;
	let hi = vec.dataLength - 1;
	let found = -1;
	const p: Probe = { index: -1, value: null };
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1;
		const j = probe(vec, order, mid, hi, p).index;
		if (j < 0) {
			hi = mid - 1;
		} else if (pred(order.compare(p.value))) {
			found = j;
			hi = mid - 1;
		} else {
			lo = j + 1;
		}
	}
	return found;
}

/**
 * The approximate match of VLOOKUP / HLOOKUP (range_lookup TRUE), MATCH
 * (match_type 1 / -1) and LOOKUP: a binary search for the last element that
 * is <= `x` (ascending data) or >= `x` (descending data). Among equal
 * values the last one wins, as in Excel. Unsorted data gives the same
 * "unexpected" answers Excel gives.
 */
export function approximateMatch(vec: Vec, x: Scalar, descending = false): number {
	const order = orderAgainst(x);
	return binarySearchLast(vec, order, descending ? (s) => s >= 0 : (s) => s <= 0);
}

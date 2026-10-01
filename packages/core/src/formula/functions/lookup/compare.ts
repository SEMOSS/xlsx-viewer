import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { compareText, toText } from '../../coerce';
import { hasWildcards, wildcardToRegExp } from '../../criteria';
import type { Scalar } from '../../types';

/**
 * Value matching and ordering for the lookup functions.
 *
 * Lookups never coerce between types: the number 1 does not match the text
 * "1" and TRUE does not match 1. Text compares case-insensitively. Blank
 * cells are never matched by a non-blank lookup value.
 */

export type Matcher = (v: Scalar) => boolean;

export interface ExactOptions {
	/** Text lookup values use `*`, `?` and `~` wildcards (VLOOKUP / MATCH exact, XLOOKUP mode 2). */
	wildcards?: boolean;
	/**
	 * A blank or empty-text lookup value also matches blank cells (XLOOKUP /
	 * XMATCH). The legacy functions only match cells holding empty text.
	 */
	blankMatchesBlank?: boolean;
}

function textMatcher(x: string, wildcards: boolean): Matcher {
	if (wildcards && hasWildcards(x)) {
		const re = wildcardToRegExp(x);
		return (v) => typeof v === 'string' && re.test(v);
	}
	const plain = wildcards ? x.replace(/~([*?~])/g, '$1') : x;
	const lower = plain.toLowerCase();
	const len = plain.length;
	return (v) =>
		typeof v === 'string' && v.length === len && (v === plain || v.toLowerCase() === lower);
}

/** Predicate for an exact match of `x` (an error lookup value matches nothing). */
export function exactMatcher(x: Scalar, opts: ExactOptions = {}): Matcher {
	if (x === null || x === '') {
		return opts.blankMatchesBlank ? (v) => v === null || v === '' : (v) => v === '';
	}
	switch (typeof x) {
		case 'number':
		case 'boolean':
			return (v) => v === x;
		case 'string':
			return textMatcher(x, opts.wildcards === true);
		default:
			return () => false;
	}
}

/**
 * Predicate for a regular-expression match (XLOOKUP / XMATCH match_mode 3):
 * the pattern may match anywhere in the text of a non-blank, non-error value.
 */
export function regexMatcher(x: Scalar): Matcher | CellError {
	const pattern = toText(x);
	if (isCellError(pattern)) {
		return pattern;
	}
	let re: RegExp;
	try {
		re = new RegExp(pattern, 'u');
	} catch {
		try {
			re = new RegExp(pattern);
		} catch {
			return ERRORS.VALUE;
		}
	}
	return (v) => {
		if (v === null || isCellError(v)) {
			return false;
		}
		return re.test(typeof v === 'string' ? v : (toText(v) as string));
	};
}

/** Ordering of cell values against a fixed lookup value, restricted to values of its own type. */
export interface Order {
	/** `v` has the lookup value's type (blanks and errors are never comparable). */
	comparable(v: Scalar): boolean;
	/** Sign of `v - x` for a comparable `v`. */
	compare(v: Scalar): number;
}

const NOTHING: Order = { comparable: () => false, compare: () => 0 };

/** Order against `x`; a blank lookup value orders as empty text. */
export function orderAgainst(x: Scalar): Order {
	const target = x === null ? '' : x;
	switch (typeof target) {
		case 'number':
			return {
				comparable: (v) => typeof v === 'number',
				compare: (v) => ((v as number) < target ? -1 : (v as number) > target ? 1 : 0),
			};
		case 'string':
			return {
				comparable: (v) => typeof v === 'string',
				compare: (v) => compareText(v as string, target),
			};
		case 'boolean':
			return {
				comparable: (v) => typeof v === 'boolean',
				compare: (v) => (v === target ? 0 : v ? 1 : -1),
			};
		default:
			return NOTHING;
	}
}

/** Compare two values of the same type (numbers, text or logicals). */
export function compareSameType(a: Scalar, b: Scalar): number {
	if (typeof a === 'number' && typeof b === 'number') {
		return a < b ? -1 : a > b ? 1 : 0;
	}
	if (typeof a === 'string' && typeof b === 'string') {
		return compareText(a, b);
	}
	if (typeof a === 'boolean' && typeof b === 'boolean') {
		return a === b ? 0 : a ? 1 : -1;
	}
	return 0;
}

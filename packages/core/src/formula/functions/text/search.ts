import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { integerArg, textArg } from '../math/args';
import { alias, escapeRegExp } from './text-utils';

const patternCache = new Map<string, RegExp>();

/**
 * SEARCH pattern: `?` matches one character, `*` any run, `~` escapes a
 * following `?`, `*` or `~`. Matching is case-insensitive and unanchored.
 */
export function searchPattern(find: string): RegExp {
	let re = patternCache.get(find);
	if (re) {
		return re;
	}
	let src = '';
	for (let i = 0; i < find.length; i++) {
		const ch = find[i];
		const next = find[i + 1];
		if (ch === '~' && (next === '*' || next === '?' || next === '~')) {
			src += escapeRegExp(next);
			i++;
		} else if (ch === '*') {
			src += '[\\s\\S]*';
		} else if (ch === '?') {
			src += '[\\s\\S]';
		} else {
			src += escapeRegExp(ch);
		}
	}
	re = new RegExp(src, 'i');
	if (patternCache.size > 500) {
		patternCache.clear();
	}
	patternCache.set(find, re);
	return re;
}

/** FIND (case-sensitive, literal) and SEARCH (case-insensitive, wildcards). */
function finder(name: string, wildcards: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 2,
		maxArgs: 3,
		params: ['value', 'value', 'value'],
		impl: (args, ctx) => {
			const find = textArg(args[0]);
			if (isCellError(find)) {
				return find;
			}
			const within = textArg(args[1]);
			if (isCellError(within)) {
				return within;
			}
			let start = 1;
			if (args.length > 2) {
				const s = integerArg(args[2], ctx);
				if (typeof s !== 'number') {
					return s;
				}
				start = s;
			}
			if (start < 1 || start > within.length) {
				return ERRORS.VALUE;
			}
			if (find === '') {
				return start;
			}
			let at: number;
			if (wildcards) {
				const found = within.slice(start - 1).search(searchPattern(find));
				at = found < 0 ? -1 : found + start - 1;
			} else {
				at = within.indexOf(find, start - 1);
			}
			return at < 0 ? ERRORS.VALUE : at + 1;
		},
	};
}

const FIND = finder('FIND', false);
const SEARCH = finder('SEARCH', true);

export const SEARCH_FUNCTIONS: FunctionDefinition[] = [
	FIND,
	alias(FIND, 'FINDB'),
	SEARCH,
	alias(SEARCH, 'SEARCHB'),
];

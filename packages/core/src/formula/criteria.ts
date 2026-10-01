import { isCellError, parseErrorCode } from '../values/errors';
import { parseNumericText } from '../values/parse-number';
import { compareText } from './coerce';
import type { Scalar } from './types';

const regexCache = new Map<string, RegExp>();

/**
 * Compile an Excel wildcard pattern (`*` any run, `?` one character, `~`
 * escapes the next character) to an anchored, case-insensitive RegExp.
 */
export function wildcardToRegExp(pattern: string): RegExp {
	const cached = regexCache.get(pattern);
	if (cached) {
		return cached;
	}
	let src = '^';
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '~' && i + 1 < pattern.length) {
			i++;
			src += pattern[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		} else if (ch === '*') {
			src += '[\\s\\S]*';
		} else if (ch === '?') {
			src += '[\\s\\S]';
		} else {
			src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		}
	}
	src += '$';
	const re = new RegExp(src, 'i');
	if (regexCache.size > 500) {
		regexCache.clear();
	}
	regexCache.set(pattern, re);
	return re;
}

/** The pattern contains unescaped wildcards. */
export function hasWildcards(pattern: string): boolean {
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '~') {
			i++;
			continue;
		}
		if (ch === '*' || ch === '?') {
			return true;
		}
	}
	return false;
}

type Op = '=' | '<>' | '<' | '>' | '<=' | '>=';

function splitOperator(text: string): { op: Op; rest: string } {
	for (const op of ['<=', '>=', '<>', '=', '<', '>'] as const) {
		if (text.startsWith(op)) {
			return { op, rest: text.slice(op.length) };
		}
	}
	return { op: '=', rest: text };
}

function cmp(op: Op, c: number): boolean {
	switch (op) {
		case '=':
			return c === 0;
		case '<>':
			return c !== 0;
		case '<':
			return c < 0;
		case '>':
			return c > 0;
		case '<=':
			return c <= 0;
		case '>=':
			return c >= 0;
	}
}

/**
 * Build a matcher for COUNTIF / SUMIF / AVERAGEIF / *IFS criteria.
 *
 * - A number matches equal numbers (and text that reads as that number).
 * - Text may start with `=`, `<>`, `<`, `>`, `<=`, `>=`. The rest is compared
 *   numerically when it reads as a number (or date / time), as a boolean
 *   when it is TRUE / FALSE, as an error when it is an error literal, and
 *   otherwise as case-insensitive text with `*` / `?` wildcards for `=` and `<>`.
 * - `"="` matches blank cells, `"<>"` matches non-blank cells, `""` matches
 *   blank cells and empty text.
 */
export function createCriteriaMatcher(criteria: Scalar, date1904 = false): (v: Scalar) => boolean {
	if (criteria === null) {
		return (v) => v === null || v === '';
	}
	if (typeof criteria === 'number') {
		return (v) => {
			if (typeof v === 'number') {
				return v === criteria;
			}
			if (typeof v === 'string') {
				const p = parseNumericText(v, { date1904 });
				return p !== null && p.value === criteria;
			}
			return false;
		};
	}
	if (typeof criteria === 'boolean') {
		return (v) => v === criteria;
	}
	if (isCellError(criteria)) {
		return (v) => isCellError(v) && v.code === criteria.code;
	}

	if (criteria === '') {
		return (v) => v === null || v === '';
	}
	const { op, rest } = splitOperator(criteria);
	if (rest === '') {
		if (op === '=') {
			return (v) => v === null;
		}
		if (op === '<>') {
			return (v) => v !== null && v !== '';
		}
		// "<" or ">" against empty text compares text only.
		return (v) => typeof v === 'string' && cmp(op, compareText(v, ''));
	}
	const num = parseNumericText(rest, { date1904 });
	if (num !== null) {
		const target = num.value;
		return (v) => {
			let n: number | null = null;
			if (typeof v === 'number') {
				n = v;
			} else if (typeof v === 'string' && (op === '=' || op === '<>')) {
				const p = parseNumericText(v, { date1904 });
				n = p ? p.value : null;
			}
			if (n === null) {
				return op === '<>';
			}
			return cmp(op, n < target ? -1 : n > target ? 1 : 0);
		};
	}

	const upper = rest.toUpperCase();
	if (upper === 'TRUE' || upper === 'FALSE') {
		const b = upper === 'TRUE';
		return (v) => {
			if (typeof v !== 'boolean') {
				return op === '<>';
			}
			return cmp(op, v === b ? 0 : v ? 1 : -1);
		};
	}

	const errCode = parseErrorCode(rest);
	if (errCode) {
		return (v) => {
			const match = isCellError(v) && v.code === errCode;
			return op === '<>' ? !match : op === '=' ? match : false;
		};
	}

	if (op === '=' || op === '<>') {
		const wild = hasWildcards(rest);
		const re = wild ? wildcardToRegExp(rest) : null;
		const plain = wild ? '' : rest.replace(/~([*?~])/g, '$1');
		return (v) => {
			let match: boolean;
			if (typeof v !== 'string') {
				match = false;
			} else if (re) {
				match = re.test(v);
			} else {
				match = compareText(v, plain) === 0;
			}
			return op === '=' ? match : !match;
		};
	}

	return (v) => typeof v === 'string' && cmp(op, compareText(v, rest));
}

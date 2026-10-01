/**
 * Format code tokenizer: splits a code into `;` sections and lexes each into
 * {@link NumberFormatToken}s. Context-dependent meaning (thousands vs scaling
 * commas, minutes vs months, fraction slashes) is resolved by the compilers.
 */
import { NAMED_COLORS } from './locale-en-us';
import type { CompareOp, DateLetter, NumberFormatToken, Placeholder, TimeUnit } from './types';

const CONDITION_RE =
	/^\s*(<>|<=|>=|=<|=>|<|>|=)\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s*$/;
const COLOR_INDEX_RE = /^color\s*(\d+)$/i;
const ELAPSED_RE = /^(?:h+|m+|s+)$/i;
const HEX_RE = /^[0-9a-f]+$/i;

/** The character at `i`: one UTF-16 unit, or a whole surrogate pair. */
function charAt(s: string, i: number): string {
	const c = s.charCodeAt(i);
	if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
		const d = s.charCodeAt(i + 1);
		if (d >= 0xdc00 && d <= 0xdfff) {
			return s.slice(i, i + 2);
		}
	}
	return s.charAt(i);
}

function normalizeOp(op: string): CompareOp {
	if (op === '=<') {
		return '<=';
	}
	if (op === '=>') {
		return '>=';
	}
	return op as CompareOp;
}

/** `[$€-407]`, `[$USD]`, `[$-409]`, `[$-F800]`, `[$-x-sysdate]` (content after the `$`). */
function localeToken(body: string, raw: string): NumberFormatToken {
	const dash = body.indexOf('-');
	const symbol = dash < 0 ? body : body.slice(0, dash);
	const tag = dash < 0 ? '' : body.slice(dash + 1);
	const lower = tag.toLowerCase();
	let lcid: number | undefined;
	let system: 'date' | 'time' | undefined;
	if (lower === 'x-sysdate') {
		system = 'date';
	} else if (lower === 'x-systime') {
		system = 'time';
	} else if (HEX_RE.test(tag)) {
		lcid = parseInt(tag, 16);
		const base = lcid & 0xffff;
		if (base === 0xf800) {
			system = 'date';
		} else if (base === 0xf400) {
			system = 'time';
		}
	}
	return { type: 'locale', symbol, lcid, system, raw };
}

function bracketToken(content: string, raw: string): NumberFormatToken {
	if (content.charAt(0) === '$') {
		return localeToken(content.slice(1), raw);
	}
	const cond = CONDITION_RE.exec(content);
	if (cond) {
		return { type: 'condition', op: normalizeOp(cond[1]), value: Number(cond[2]), raw };
	}
	const named = NAMED_COLORS.get(content.toLowerCase());
	if (named) {
		return { type: 'color', hex: named, raw };
	}
	const indexed = COLOR_INDEX_RE.exec(content);
	if (indexed) {
		const index = Number(indexed[1]);
		return index >= 1 && index <= 56 ? { type: 'color', index, raw } : { type: 'ignored', raw };
	}
	if (ELAPSED_RE.test(content)) {
		const unit = content.charAt(0).toLowerCase() as TimeUnit;
		return { type: 'elapsed', unit, length: content.length, raw };
	}
	return { type: 'ignored', raw };
}

/** A run of the same date letter (case-insensitive): `yyyy`, `mmm`, `dd`. */
function dateRun(code: string, i: number, letter: DateLetter, out: NumberFormatToken[]): number {
	let j = i + 1;
	while (j < code.length && code.charAt(j).toLowerCase() === letter) {
		j++;
	}
	out.push({ type: 'date', part: letter, length: j - i, raw: code.slice(i, j) });
	return j;
}

/** Letters, numerals and other characters; returns the next index. */
function wordToken(code: string, i: number, out: NumberFormatToken[]): number {
	const c = code.charAt(i);
	const cc = code.charCodeAt(i);
	if (cc >= 49 && cc <= 57) {
		out.push({ type: 'numeral', text: c, raw: c });
		return i + 1;
	}
	const lower = c.toLowerCase();
	if (c === 't' && out.length === 0 && /[0#?]/.test(code.charAt(i + 1))) {
		// Thai builtins (ids 59-70: `t0.00`, `t# ?/?`): `t` selects Thai digits.
		out.push({ type: 'ignored', raw: c });
		return i + 1;
	}
	switch (lower) {
		case 'e': {
			const next = code.charAt(i + 1);
			if (next === '+' || next === '-') {
				out.push({ type: 'exponent', upper: c === 'E', plus: next === '+', raw: c + next });
				return i + 2;
			}
			return dateRun(code, i, 'e', out);
		}
		case 'g':
			if (code.slice(i, i + 7).toLowerCase() === 'general') {
				out.push({ type: 'general', raw: code.slice(i, i + 7) });
				return i + 7;
			}
			return dateRun(code, i, 'g', out);
		case 'b': {
			const next = code.charAt(i + 1);
			if (next === '1' || next === '2') {
				// B1 / B2 calendar selectors (Gregorian / Hijri): no en-US output.
				out.push({ type: 'ignored', raw: c + next });
				return i + 2;
			}
			return dateRun(code, i, 'b', out);
		}
		case 'a': {
			const long = code.slice(i, i + 5);
			if (long.toUpperCase() === 'AM/PM') {
				const am = long.charAt(0) === 'A' ? 'AM' : 'am';
				const pm = long.charAt(3) === 'P' ? 'PM' : 'pm';
				out.push({ type: 'ampm', am, pm, raw: long });
				return i + 5;
			}
			const short = code.slice(i, i + 3);
			if (short.toUpperCase() === 'A/P') {
				out.push({ type: 'ampm', am: short.charAt(0), pm: short.charAt(2), raw: short });
				return i + 3;
			}
			break;
		}
		case 'y':
		case 'm':
		case 'd':
		case 'h':
		case 's':
			return dateRun(code, i, lower, out);
		default:
			break;
	}
	const ch = charAt(code, i);
	out.push({ type: 'literal', text: ch, raw: ch });
	return i + ch.length;
}

/**
 * Tokenize a format code into its sections (split at unquoted `;`). Every
 * character is covered by exactly one token, so `raw` values round-trip.
 */
export function tokenizeFormat(code: string): NumberFormatToken[][] {
	const sections: NumberFormatToken[][] = [];
	let tokens: NumberFormatToken[] = [];
	const n = code.length;
	let i = 0;
	while (i < n) {
		const c = code.charAt(i);
		switch (c) {
			case ';':
				sections.push(tokens);
				tokens = [];
				i++;
				break;
			case '"': {
				let j = code.indexOf('"', i + 1);
				if (j < 0) {
					j = n;
				}
				tokens.push({ type: 'literal', text: code.slice(i + 1, j), raw: code.slice(i, j + 1) });
				i = j + 1;
				break;
			}
			case '\\':
			case '_':
			case '*': {
				if (i + 1 >= n) {
					tokens.push({ type: 'ignored', raw: c });
					i++;
					break;
				}
				const ch = charAt(code, i + 1);
				const raw = c + ch;
				if (c === '\\') {
					tokens.push({ type: 'literal', text: ch, raw });
				} else if (c === '_') {
					tokens.push({ type: 'space', char: ch, raw });
				} else {
					tokens.push({ type: 'fill', char: ch, raw });
				}
				i += raw.length;
				break;
			}
			case '[': {
				const j = code.indexOf(']', i + 1);
				if (j < 0) {
					tokens.push({ type: 'literal', text: c, raw: c });
					i++;
				} else {
					tokens.push(bracketToken(code.slice(i + 1, j), code.slice(i, j + 1)));
					i = j + 1;
				}
				break;
			}
			case '0':
			case '#':
			case '?':
				tokens.push({ type: 'digit', char: c as Placeholder, raw: c });
				i++;
				break;
			case '.':
				tokens.push({ type: 'point', raw: c });
				i++;
				break;
			case ',':
				tokens.push({ type: 'comma', raw: c });
				i++;
				break;
			case '%':
				tokens.push({ type: 'percent', raw: c });
				i++;
				break;
			case '/':
				tokens.push({ type: 'slash', raw: c });
				i++;
				break;
			case '@':
				tokens.push({ type: 'text', raw: c });
				i++;
				break;
			default:
				i = wordToken(code, i, tokens);
		}
	}
	sections.push(tokens);
	return sections;
}

import { numberToText } from '../formula/coerce';
import type { CellValue } from '../types/cell';

/** Rule types of the "Text that Contains" family. */
export type TextRuleKind = 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';

/** Text a cell value converts to in formulas (numbers in General form); `undefined` for errors. */
export function cellValueText(value: CellValue): string | undefined {
	if (value === null) {
		return '';
	}
	if (typeof value === 'string') {
		return value;
	}
	if (typeof value === 'number') {
		return numberToText(value);
	}
	if (typeof value === 'boolean') {
		return value ? 'TRUE' : 'FALSE';
	}
	return undefined;
}

/**
 * Blank test of "Format only cells that contain: Blanks", equivalent to
 * Excel's `LEN(TRIM(A1))=0`: empty cells and text made only of spaces are
 * blank. Errors are neither blank nor non-blank (`undefined`).
 */
export function isBlankForRule(value: CellValue): boolean | undefined {
	if (value === null) {
		return true;
	}
	if (typeof value === 'string') {
		return /^ *$/.test(value);
	}
	if (typeof value === 'object') {
		return undefined;
	}
	return false;
}

function escapeRegExp(ch: string): string {
	return /[\\^$.*+?()[\]{}|/]/.test(ch) ? `\\${ch}` : ch;
}

/**
 * Case-insensitive substring test with SEARCH wildcards: `?` matches one
 * character, `*` any run, `~` escapes the next wildcard.
 */
export function createSearchMatcher(pattern: string): (text: string) => boolean {
	if (!/[*?~]/.test(pattern)) {
		const lower = pattern.toLowerCase();
		return (text) => text.toLowerCase().includes(lower);
	}
	let source = '';
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '~' && i + 1 < pattern.length && /[*?~]/.test(pattern[i + 1])) {
			source += escapeRegExp(pattern[i + 1]);
			i++;
		} else if (ch === '*') {
			source += '[\\s\\S]*';
		} else if (ch === '?') {
			source += '[\\s\\S]';
		} else {
			source += escapeRegExp(ch);
		}
	}
	const re = new RegExp(source, 'i');
	return (text) => re.test(text);
}

/**
 * Matcher for a text rule, following the formulas Excel writes for them:
 * contains = `NOT(ISERROR(SEARCH(t,A1)))` (wildcards allowed; errors never
 * contain, so they always "do not contain"), begins / ends with =
 * `LEFT(A1,LEN(t))=t` / `RIGHT(...)` (literal, errors never match). All
 * comparisons are case-insensitive; numbers compare by their General text.
 */
export function createTextMatcher(kind: TextRuleKind, text: string): (value: CellValue) => boolean {
	const needle = text.toLowerCase();
	switch (kind) {
		case 'containsText': {
			const search = createSearchMatcher(text);
			return (value) => {
				const s = cellValueText(value);
				return s !== undefined && search(s);
			};
		}
		case 'notContainsText': {
			const search = createSearchMatcher(text);
			return (value) => {
				const s = cellValueText(value);
				return s === undefined || !search(s);
			};
		}
		case 'beginsWith':
			return (value) => {
				const s = cellValueText(value);
				return s !== undefined && s.toLowerCase().startsWith(needle);
			};
		case 'endsWith':
			return (value) => {
				const s = cellValueText(value);
				return s !== undefined && s.toLowerCase().endsWith(needle);
			};
		default:
			return () => false;
	}
}

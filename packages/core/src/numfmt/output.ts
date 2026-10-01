/**
 * Shared building blocks for compiling and rendering sections: the render
 * item list builder, literal text of tokens, the output buffer (with the
 * repeat-fill split) and digit placeholder emission.
 */
import type { NumberFormatToken, FormattedValue, Placeholder, RenderItem } from './types';

/** Accumulates render items, merging adjacent literals. */
export class ItemList {
	readonly items: RenderItem[] = [];

	lit(text: string): void {
		if (text === '') {
			return;
		}
		const last = this.items.length - 1;
		const prev = this.items[last];
		if (prev !== undefined && prev.k === 'lit') {
			this.items[last] = { k: 'lit', text: prev.text + text };
		} else {
			this.items.push({ k: 'lit', text });
		}
	}

	push(item: RenderItem): void {
		this.items.push(item);
	}
}

/** Display text of a token rendered literally (outside its special role). */
export function literalText(t: NumberFormatToken): string {
	switch (t.type) {
		case 'literal':
		case 'numeral':
			return t.text;
		case 'space':
			return ' ';
		case 'point':
			return '.';
		case 'comma':
			return ',';
		case 'percent':
			return '%';
		case 'slash':
			return '/';
		case 'text':
		case 'fill':
		case 'color':
		case 'condition':
		case 'locale':
		case 'ignored':
			return '';
		default:
			return t.raw;
	}
}

/** Output under construction. Text before the first `*x` fill moves to `before`. */
export interface Out {
	s: string;
	before: string;
	fill: string | null;
}

export function newOut(): Out {
	return { s: '', before: '', fill: null };
}

/** Start the repeat-fill region (only the first fill of a section counts). */
export function putFill(out: Out, char: string): void {
	if (out.fill === null) {
		out.fill = char;
		out.before = out.s;
		out.s = '';
	}
}

/** Final result with an optional leading sign and color. */
export function finish(out: Out, sign: string, color: string | undefined): FormattedValue {
	let result: FormattedValue;
	if (out.fill === null) {
		result = { text: sign + out.s };
	} else {
		const before = sign + out.before;
		result = { text: before + out.s, fill: { before, char: out.fill, after: out.s } };
	}
	if (color !== undefined) {
		result.color = color;
	}
	return result;
}

/** What an unused placeholder shows: `0` a zero, `?` a space, `#` nothing. */
export function padFor(ph: Placeholder): string {
	return ph === '0' ? '0' : ph === '?' ? ' ' : '';
}

/** Insert thousands separators into an integer digit string. */
export function groupDigits(digits: string): string {
	const n = digits.length;
	if (n <= 3) {
		return digits;
	}
	let s = '';
	for (let i = 0; i < n; i++) {
		s += digits.charAt(i);
		const p = n - 1 - i;
		if (p > 0 && p % 3 === 0) {
			s += ',';
		}
	}
	return s;
}

/**
 * Text for an integer placeholder at `pos` (0 = units) of a run of `slots`
 * placeholders. The leftmost placeholder also prints any digits that do not
 * fit; with `grouping` a separator follows each digit at a thousands position.
 */
export function intSlot(
	digits: string,
	slots: number,
	pos: number,
	ph: Placeholder,
	grouping: boolean,
): string {
	const len = digits.length;
	let s = '';
	if (pos === slots - 1 && len > slots) {
		for (let p = len - 1; p >= slots; p--) {
			s += digits.charAt(len - 1 - p);
			if (grouping && p % 3 === 0) {
				s += ',';
			}
		}
	}
	const sep = grouping && pos > 0 && pos % 3 === 0;
	if (pos < len) {
		s += digits.charAt(len - 1 - pos);
		if (sep) {
			s += ',';
		}
	} else if (ph === '0') {
		s += sep ? '0,' : '0';
	} else if (ph === '?') {
		s += sep ? '  ' : ' ';
	}
	return s;
}

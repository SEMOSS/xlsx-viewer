/**
 * Compilers for numeric sections: plain numbers and scientific notation,
 * fractions, and the simple General / text sections.
 */
import { ItemList, literalText } from './output';
import type { NumberFormatToken, RenderItem } from './types';

export interface NumberLayout {
	items: RenderItem[];
	intSlots: number;
	decSlots: number;
	expSlots: number;
	grouping: boolean;
	scale: number;
	percent: boolean;
}

export interface FractionLayout {
	items: RenderItem[];
	intSlots: number;
	numSlots: number;
	denSlots: number;
	denFixed: number;
	grouping: boolean;
	scale: number;
	percent: boolean;
}

type CommaRole = 'group' | 'scale' | 'ignore' | 'literal';

/**
 * Excel: a comma enclosed by digit placeholders of the integer part is the
 * thousands separator; a comma that follows the last placeholder (of the
 * integer or decimal part) scales by 1,000; other commas are literal.
 */
function commaRole(
	body: readonly NumberFormatToken[],
	i: number,
	intEnd: number,
	mantEnd: number,
	firstInt: number,
	lastInt: number,
): CommaRole {
	let p = i - 1;
	while (p >= 0 && body[p].type === 'comma') {
		p--;
	}
	let q = i + 1;
	while (q < mantEnd && body[q].type === 'comma') {
		q++;
	}
	const prevDigit = p >= 0 && (body[p].type === 'digit' || body[p].type === 'point');
	const nextDigit = q < mantEnd && body[q].type === 'digit';
	if (i < intEnd) {
		if (firstInt >= 0 && firstInt < i && lastInt > i) {
			return 'group';
		}
		return prevDigit ? 'scale' : 'literal';
	}
	if (prevDigit) {
		return nextDigit ? 'ignore' : 'scale';
	}
	return 'literal';
}

/** Plain number (`#,##0.00`, `0%`, `(###) ###-####`) or scientific (`0.00E+00`) section. */
export function compileNumber(
	body: readonly NumberFormatToken[],
	scientific: boolean,
): NumberLayout {
	const n = body.length;
	let expIdx = -1;
	if (scientific) {
		expIdx = body.findIndex((t) => t.type === 'exponent');
	}
	const mantEnd = expIdx < 0 ? n : expIdx;
	let pointIdx = -1;
	for (let i = 0; i < mantEnd; i++) {
		if (body[i].type === 'point') {
			pointIdx = i;
			break;
		}
	}
	const intEnd = pointIdx < 0 ? mantEnd : pointIdx;
	let intSlots = 0;
	let decSlots = 0;
	let expSlots = 0;
	let firstInt = -1;
	let lastInt = -1;
	for (let i = 0; i < n; i++) {
		if (body[i].type !== 'digit') {
			continue;
		}
		if (i < intEnd) {
			intSlots++;
			if (firstInt < 0) {
				firstInt = i;
			}
			lastInt = i;
		} else if (i < mantEnd) {
			decSlots++;
		} else {
			expSlots++;
		}
	}

	const list = new ItemList();
	let grouping = false;
	let scale = 0;
	let percent = false;
	let intPos = intSlots;
	let decIdx = 0;
	let expPos = expSlots;
	for (let i = 0; i < n; i++) {
		const t = body[i];
		switch (t.type) {
			case 'digit':
				if (i < intEnd) {
					list.push({ k: 'int', pos: --intPos, ph: t.char });
				} else if (i < mantEnd) {
					list.push({ k: 'dec', idx: decIdx++, ph: t.char });
				} else {
					list.push({ k: 'exp', pos: --expPos, ph: t.char });
				}
				break;
			case 'point':
				if (i === pointIdx) {
					if (intSlots === 0) {
						list.push({ k: 'intAll' });
					}
					list.push({ k: 'point' });
				} else {
					list.lit('.');
				}
				break;
			case 'comma': {
				const role =
					i < mantEnd ? commaRole(body, i, intEnd, mantEnd, firstInt, lastInt) : 'literal';
				if (role === 'group') {
					grouping = true;
				} else if (role === 'scale') {
					scale -= 3;
				} else if (role === 'literal') {
					list.lit(',');
				}
				break;
			}
			case 'percent':
				scale += 2;
				percent = true;
				list.lit('%');
				break;
			case 'exponent':
				if (i === expIdx) {
					list.push({ k: 'expMark', upper: t.upper, plus: t.plus });
				} else {
					list.lit(t.raw);
				}
				break;
			case 'fill':
				list.push({ k: 'fill', char: t.char });
				break;
			default:
				list.lit(literalText(t));
		}
	}
	return { items: list.items, intSlots, decSlots, expSlots, grouping, scale, percent };
}

function isBlank(t: NumberFormatToken | undefined): boolean {
	return t !== undefined && t.type === 'literal' && t.text !== '' && t.text.trim() === '';
}

/**
 * Index of a fraction slash: a `/` with a digit placeholder before it and a
 * placeholder or a fixed denominator after it (spaces allowed around it).
 */
export function findFractionSlash(body: readonly NumberFormatToken[]): number {
	for (let i = 0; i < body.length; i++) {
		if (body[i].type !== 'slash') {
			continue;
		}
		let a = i - 1;
		while (isBlank(body[a])) {
			a--;
		}
		let b = i + 1;
		while (isBlank(body[b])) {
			b++;
		}
		const before = body[a];
		const after = body[b];
		if (
			before !== undefined &&
			before.type === 'digit' &&
			after !== undefined &&
			(after.type === 'digit' || after.type === 'numeral')
		) {
			return i;
		}
	}
	return -1;
}

/** Fraction section (`# ?/?`, `# ??/??`, `?/?`, `# ?/8`, `# ??/100`). */
export function compileFraction(body: readonly NumberFormatToken[], slash: number): FractionLayout {
	const n = body.length;
	let a = slash - 1;
	while (isBlank(body[a])) {
		a--;
	}
	let numStart = a;
	while (numStart > 0 && body[numStart - 1].type === 'digit') {
		numStart--;
	}
	let b = slash + 1;
	while (isBlank(body[b])) {
		b++;
	}
	let denEnd = b;
	let denText = '';
	let denSlots = 0;
	if (body[b].type === 'numeral') {
		for (; denEnd < n; denEnd++) {
			const t = body[denEnd];
			if (t.type !== 'numeral' && !(t.type === 'digit' && t.char === '0')) {
				break;
			}
			denText += t.raw;
		}
	} else {
		while (denEnd < n && body[denEnd].type === 'digit') {
			denSlots++;
			denEnd++;
		}
	}
	let intSlots = 0;
	let firstWhole = -1;
	let lastWhole = -1;
	for (let i = 0; i < numStart; i++) {
		if (body[i].type === 'digit') {
			intSlots++;
			if (firstWhole < 0) {
				firstWhole = i;
			}
			lastWhole = i;
		}
	}
	const numSlots = a - numStart + 1;

	const list = new ItemList();
	let grouping = false;
	let scale = 0;
	let percent = false;
	let wholePos = intSlots;
	let numPos = numSlots;
	let denIdx = 0;
	for (let i = 0; i < n; i++) {
		const t = body[i];
		if (i === slash) {
			list.push({ k: 'slash' });
		} else if (i >= numStart && i <= a && t.type === 'digit') {
			list.push({ k: 'num', pos: --numPos, ph: t.char });
		} else if (i >= b && i < denEnd) {
			if (denText !== '') {
				if (i === b) {
					list.push({ k: 'denFixed', text: denText });
				}
			} else if (t.type === 'digit') {
				list.push({ k: 'den', idx: denIdx++, ph: t.char });
			}
		} else if (t.type === 'digit' && i < numStart) {
			list.push({ k: 'whole', pos: --wholePos, ph: t.char });
		} else if (t.type === 'comma' && i < numStart && firstWhole < i && lastWhole > i) {
			grouping = true;
		} else if (t.type === 'percent') {
			scale += 2;
			percent = true;
			list.lit('%');
		} else if (t.type === 'fill') {
			list.push({ k: 'fill', char: t.char });
		} else {
			list.lit(literalText(t));
		}
	}
	const denFixed = denText === '' ? 0 : Number(denText);
	return { items: list.items, intSlots, numSlots, denSlots, denFixed, grouping, scale, percent };
}

/**
 * General and text sections: literals, fills and the value itself (`General`
 * or `@`). In a text section a `General` token also shows the text.
 */
export function compileSimple(body: readonly NumberFormatToken[], text: boolean): RenderItem[] {
	const list = new ItemList();
	for (const t of body) {
		if (t.type === 'fill') {
			list.push({ k: 'fill', char: t.char });
		} else if (t.type === 'general') {
			list.push(text ? { k: 'text' } : { k: 'general' });
		} else if (t.type === 'text') {
			if (text) {
				list.push({ k: 'text' });
			}
		} else {
			list.lit(literalText(t));
		}
	}
	return list.items;
}

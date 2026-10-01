/**
 * Renderers for plain number, scientific, General and text sections.
 */
import { roundDecimal, shiftDecimal, splitDecimal, toDecimal, ZERO, type Decimal } from './decimal';
import { formatGeneral } from './general';
import { groupDigits, intSlot, padFor, putFill, type Out } from './output';
import type { NumberFormatSection } from './types';

/**
 * Exponent for scientific notation. One integer placeholder gives a mantissa
 * in [1, 10); n > 1 placeholders make the exponent a multiple of n
 * (engineering notation for `##0.0E+0`); none gives a mantissa in [0.1, 1).
 */
function exponentFor(magnitude: number, intSlots: number): number {
	if (intSlots <= 1) {
		return intSlots === 0 ? magnitude + 1 : magnitude;
	}
	return Math.floor(magnitude / intSlots) * intSlots;
}

/** Rounded mantissa and exponent; a carry (9.99 -> 10.0) moves to the next exponent. */
function scientificParts(
	d: Decimal,
	intSlots: number,
	decSlots: number,
): { mantissa: Decimal; exponent: number } {
	if (d.digits === '') {
		return { mantissa: ZERO, exponent: 0 };
	}
	let exponent = exponentFor(d.exp, intSlots);
	let mantissa = roundDecimal({ digits: d.digits, exp: d.exp - exponent }, decSlots);
	if (mantissa.digits === '') {
		return { mantissa: ZERO, exponent: 0 };
	}
	const magnitude = mantissa.exp + exponent;
	const adjusted = exponentFor(magnitude, intSlots);
	if (adjusted !== exponent) {
		mantissa = roundDecimal({ digits: mantissa.digits, exp: magnitude - adjusted }, decSlots);
		exponent = adjusted;
	}
	return { mantissa, exponent };
}

/** Render a number or scientific section for a non-negative value. */
export function renderNumber(section: NumberFormatSection, abs: number, out: Out): void {
	const { decSlots, intSlots, grouping } = section;
	let d = shiftDecimal(toDecimal(abs), section.scale);
	let exponent = 0;
	if (section.kind === 'scientific') {
		const parts = scientificParts(d, intSlots, decSlots);
		d = parts.mantissa;
		exponent = parts.exponent;
	} else {
		d = roundDecimal(d, decSlots);
	}
	const { int, frac } = splitDecimal(d, decSlots);
	let lastSignificant = frac.length - 1;
	while (lastSignificant >= 0 && frac.charCodeAt(lastSignificant) === 48) {
		lastSignificant--;
	}
	const expDigits = String(Math.abs(exponent));
	for (const item of section.items) {
		switch (item.k) {
			case 'lit':
				out.s += item.text;
				break;
			case 'fill':
				putFill(out, item.char);
				break;
			case 'int':
				out.s += intSlot(int, intSlots, item.pos, item.ph, grouping);
				break;
			case 'intAll':
				out.s += grouping ? groupDigits(int) : int;
				break;
			case 'point':
				out.s += '.';
				break;
			case 'dec':
				out.s += item.idx <= lastSignificant ? frac.charAt(item.idx) : padFor(item.ph);
				break;
			case 'expMark':
				out.s += (item.upper ? 'E' : 'e') + (exponent < 0 ? '-' : item.plus ? '+' : '');
				break;
			case 'exp':
				out.s += intSlot(expDigits, section.expSlots, item.pos, item.ph, false);
				break;
			default:
				break;
		}
	}
}

/** Render a section containing `General` (with optional literals) for a non-negative value. */
export function renderGeneralSection(
	section: NumberFormatSection,
	abs: number,
	maxChars: number,
	out: Out,
): void {
	for (const item of section.items) {
		if (item.k === 'lit') {
			out.s += item.text;
		} else if (item.k === 'fill') {
			putFill(out, item.char);
		} else if (item.k === 'general') {
			out.s += formatGeneral(abs, maxChars);
		}
	}
}

/** Render the text section for a text value. */
export function renderTextSection(section: NumberFormatSection, text: string, out: Out): void {
	for (const item of section.items) {
		if (item.k === 'lit') {
			out.s += item.text;
		} else if (item.k === 'fill') {
			putFill(out, item.char);
		} else if (item.k === 'text') {
			out.s += text;
		}
	}
}

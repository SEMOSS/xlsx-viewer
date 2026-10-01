/**
 * Format code inspection: date / time detection, the Number Format category
 * (as Excel's ribbon dropdown and Format Cells dialog would show it) and the
 * number of displayed decimals.
 */
import { CURRENCY_SYMBOLS } from './locale-en-us';
import { parseFormatCode } from './parse';
import type { FormatCategory, NumberFormatSection, NumberFormatToken } from './types';

/** Excel's en-US "Special" formats. */
export const SPECIAL_FORMAT_CODES: ReadonlySet<string> = new Set([
	'00000',
	'00000-0000',
	'[<=9999999]###-####;(###) ###-####',
	'000-00-0000',
]);

/** Literal characters that keep a number format in its standard category. */
const NUMBER_PUNCTUATION = /^[\s()+-]*$/;
/** Literal characters of standard date / time formats. */
const DATE_PUNCTUATION = /^[\s/\-.,:]*$/;

/** True for date and time formats (quoted text, escapes and color / locale tags are ignored). */
export function isDateFormat(formatCode: string): boolean {
	return parseFormatCode(formatCode).isDate;
}

/** True for formats that show only time fields (`h:mm`, `[h]:mm:ss`, `mm:ss.0`). */
export function isTimeOnlyFormat(formatCode: string): boolean {
	return parseFormatCode(formatCode).isTimeOnly;
}

function isCurrencyText(text: string): boolean {
	const t = text.trim();
	return t !== '' && (CURRENCY_SYMBOLS.has(t) || /^[A-Z]{3}$/.test(t));
}

interface LiteralScan {
	currency: boolean;
	/** Literal text other than punctuation and currency symbols. */
	other: boolean;
}

function scanLiterals(tokens: readonly NumberFormatToken[], punctuation: RegExp): LiteralScan {
	let currency = false;
	let other = false;
	for (const t of tokens) {
		if (t.type === 'locale') {
			currency ||= t.symbol !== '';
		} else if (t.type === 'literal') {
			if (isCurrencyText(t.text)) {
				currency = true;
			} else if (!punctuation.test(t.text)) {
				other = true;
			}
		}
	}
	return { currency, other };
}

function numberCategory(numeric: readonly NumberFormatSection[]): FormatCategory {
	let currency = false;
	for (const s of numeric) {
		if (s.kind === 'empty') {
			continue;
		}
		if (s.kind !== numeric[0].kind || s.scale !== numeric[0].scale) {
			return 'custom';
		}
		const scan = scanLiterals(s.tokens, NUMBER_PUNCTUATION);
		if (scan.other) {
			return 'custom';
		}
		currency ||= scan.currency;
	}
	const first = numeric[0];
	switch (first.kind) {
		case 'scientific':
			return 'scientific';
		case 'fraction':
			return 'fraction';
		case 'number':
			break;
		default:
			return 'custom';
	}
	if (first.percent) {
		return first.scale === 2 ? 'percentage' : 'custom';
	}
	if (first.scale !== 0) {
		return 'custom';
	}
	if (numeric.some((s) => s.hasFill)) {
		return 'accounting';
	}
	return currency ? 'currency' : 'number';
}

/** Broad category of a format code (drives the Number Format dropdown). */
export function classifyFormat(formatCode: string): FormatCategory {
	const parsed = parseFormatCode(formatCode);
	if (parsed.isGeneral) {
		return 'general';
	}
	if (formatCode === '@') {
		return 'text';
	}
	if (SPECIAL_FORMAT_CODES.has(formatCode)) {
		return 'special';
	}
	const first = parsed.numeric[0];
	if (
		first === undefined ||
		parsed.conditional ||
		first.color !== undefined ||
		first.colorIndex !== undefined
	) {
		return 'custom';
	}
	if (first.kind === 'date') {
		const dateSections = parsed.numeric.filter((s) => s.kind !== 'empty');
		if (dateSections.some((s) => s.kind !== 'date')) {
			return 'custom';
		}
		if (dateSections.some((s) => scanLiterals(s.tokens, DATE_PUNCTUATION).other)) {
			return 'custom';
		}
		return first.hasDate ? 'date' : 'time';
	}
	return numberCategory(parsed.numeric);
}

/** Decimal places shown by the first numeric section (fractional seconds for times). */
export function getFormatDecimals(formatCode: string): number {
	const first = parseFormatCode(formatCode).numeric[0];
	if (first === undefined) {
		return 0;
	}
	if (first.kind === 'number' || first.kind === 'scientific') {
		return first.decSlots;
	}
	return first.kind === 'date' ? first.subsecDigits : 0;
}

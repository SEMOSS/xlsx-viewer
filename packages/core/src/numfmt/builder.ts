/**
 * Format code builder for the Format Cells dialog: produces the code Excel's
 * dialog writes for a category and its options.
 */
import { MAX_FORMAT_DECIMALS } from './adjust-decimals';
import { CURRENCY_SYMBOL_PRESETS, type NegativeNumberStyle } from './presets';
import type { FormatCategory } from './types';

export interface NumberFormatSpec {
	category: FormatCategory;
	/** Decimal places (number, currency, accounting, percentage, scientific). Default 2. */
	decimals?: number;
	/** Thousands separator for the Number category (currency and accounting always group). */
	useThousands?: boolean;
	/** Negative numbers (Number and Currency). Default `'minus'`. */
	negativeStyle?: NegativeNumberStyle;
	/**
	 * Currency / accounting symbol: a `CURRENCY_SYMBOL_PRESETS` id (`'$'`, `'€'`,
	 * `'€-after'`, `'£'`, ...), `'none'` / `''` for no symbol, or any other text.
	 * Default `'$'`.
	 */
	currencySymbol?: string;
	/** Placement of a custom (non-preset) symbol. Default `'before'`. */
	symbolPosition?: 'before' | 'after';
	/** Fraction "up to N digits" (1-3 in the dialog). Default 1. */
	fractionDigits?: number;
	/** Fixed fraction denominator (2, 4, 8, 16, 10, 100); overrides `fractionDigits`. */
	fractionDenominator?: number;
	/** Date / time / special / custom code (e.g. one of the preset codes). */
	code?: string;
}

function clampDecimals(decimals: number | undefined): number {
	const d = decimals === undefined ? 2 : Math.trunc(decimals);
	return Math.max(0, Math.min(MAX_FORMAT_DECIMALS, Number.isFinite(d) ? d : 2));
}

function decimalPart(decimals: number): string {
	return decimals > 0 ? '.' + '0'.repeat(decimals) : '';
}

/** Quote a custom currency symbol for use in a format code. */
function quoteSymbol(symbol: string): string {
	return `"${symbol.replace(/"/g, '')}"`;
}

function currencyAffixes(spec: NumberFormatSpec): { prefix: string; suffix: string } {
	const symbol = spec.currencySymbol ?? '$';
	if (symbol === '') {
		return { prefix: '', suffix: '' };
	}
	const preset = CURRENCY_SYMBOL_PRESETS.find((p) => p.id === symbol);
	if (preset) {
		return { prefix: preset.prefix, suffix: preset.suffix };
	}
	return spec.symbolPosition === 'after'
		? { prefix: '', suffix: '\\ ' + quoteSymbol(symbol) }
		: { prefix: quoteSymbol(symbol), suffix: '' };
}

/** Apply a negative number style to a positive-section code. */
function withNegativeStyle(core: string, style: NegativeNumberStyle | undefined): string {
	switch (style ?? 'minus') {
		case 'red':
			return `${core};[Red]${core}`;
		case 'parens':
			return `${core}_);\\(${core}\\)`;
		case 'redParens':
			return `${core}_);[Red]\\(${core}\\)`;
		case 'redMinus':
			return `${core}_ ;[Red]\\-${core}\\ `;
		default:
			return core;
	}
}

function accountingFormat(decimals: number, prefix: string, suffix: string): string {
	const number = '#,##0' + decimalPart(decimals);
	const zero = '"-"' + '?'.repeat(decimals);
	if (suffix === '') {
		return `_(${prefix}* ${number}_);_(${prefix}* \\(${number}\\);_(${prefix}* ${zero}_);_(@_)`;
	}
	return `_(* ${number}${suffix}_);_(* \\(${number}\\)${suffix};_(* ${zero}${suffix}_);_(@_)`;
}

function fractionFormat(spec: NumberFormatSpec): string {
	const den = spec.fractionDenominator;
	if (den !== undefined && Number.isInteger(den) && den >= 2) {
		// One `?` per digit of the largest numerator (`# ?/10`, `# ??/16`, `# ??/100`).
		return `# ${'?'.repeat(String(den - 1).length)}/${den}`;
	}
	const digits = Math.max(1, Math.min(5, Math.trunc(spec.fractionDigits ?? 1)));
	const q = '?'.repeat(digits);
	return `# ${q}/${q}`;
}

/** The format code Excel's Format Cells dialog produces for `spec`. */
export function buildNumberFormat(spec: NumberFormatSpec): string {
	const decimals = clampDecimals(spec.decimals);
	switch (spec.category) {
		case 'number': {
			const core = (spec.useThousands ? '#,##0' : '0') + decimalPart(decimals);
			return withNegativeStyle(core, spec.negativeStyle);
		}
		case 'currency': {
			const { prefix, suffix } = currencyAffixes(spec);
			const core = `${prefix}#,##0${decimalPart(decimals)}${suffix}`;
			return withNegativeStyle(core, spec.negativeStyle);
		}
		case 'accounting': {
			const { prefix, suffix } = currencyAffixes(spec);
			return accountingFormat(decimals, prefix, suffix);
		}
		case 'percentage':
			return `0${decimalPart(decimals)}%`;
		case 'scientific':
			return `0${decimalPart(decimals)}E+00`;
		case 'fraction':
			return fractionFormat(spec);
		case 'date':
			return spec.code ?? 'm/d/yyyy';
		case 'time':
			return spec.code ?? '[$-F400]h:mm:ss\\ AM/PM';
		case 'text':
			return '@';
		case 'special':
			return spec.code ?? '00000';
		case 'custom':
			return spec.code ?? 'General';
		default:
			return 'General';
	}
}

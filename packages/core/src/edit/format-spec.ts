/**
 * Format Cells > Number: map an existing format code back to the dialog's
 * category and options (the inverse of `buildNumberFormat`), and list the
 * codes Excel offers in the Custom category.
 */
import { MAX_FORMAT_DECIMALS } from '../numfmt/adjust-decimals';
import { buildNumberFormat, type NumberFormatSpec } from '../numfmt/builder';
import { classifyFormat, getFormatDecimals } from '../numfmt/classify';
import {
	CURRENCY_SYMBOL_PRESETS,
	DATE_FORMAT_PRESETS,
	FRACTION_FORMAT_PRESETS,
	NEGATIVE_NUMBER_STYLES,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from '../numfmt/presets';
import type { FormatCategory } from '../numfmt/types';
import { BUILTIN_NUMBER_FORMATS } from '../styles/builtin-number-formats';
import type { StyleSheet } from '../types/style';

/** The Number tab state for a format code. */
export interface NumberFormatChoice extends NumberFormatSpec {
	category: FormatCategory;
	/**
	 * For date, time, fraction, special and custom: the code selected in the
	 * type list (a preset's code when one matches, else the code itself).
	 */
	code?: string;
}

/**
 * Comparison key for format codes: the en-US locale tag `[$-409]`, a
 * trailing text section `;@`, backslash escapes and literal quotes do not
 * change what Excel displays, so `"$"#,##0` equals `\$#,##0` and
 * `[$-409]h:mm AM/PM;@` equals `h:mm AM/PM`.
 */
export function formatCodeKey(code: string): string {
	return code
		.replace(/\[\$-(?:409|en-US)\]/gi, '')
		.replace(/;@$/, '')
		.replace(/\\(.)/g, '$1')
		.replace(/"([^"]*)"/g, '$1');
}

/** The two codes display values the same way (see {@link formatCodeKey}). */
export function sameFormatCode(a: string, b: string): boolean {
	return a === b || formatCodeKey(a) === formatCodeKey(b);
}

function clampDecimals(code: string): number {
	return Math.max(0, Math.min(MAX_FORMAT_DECIMALS, getFormatDecimals(code)));
}

function firstMatch(code: string, specs: Iterable<NumberFormatSpec>): NumberFormatSpec | undefined {
	for (const spec of specs) {
		if (sameFormatCode(buildNumberFormat(spec), code)) {
			return spec;
		}
	}
	return undefined;
}

function* numberSpecs(decimals: number): Generator<NumberFormatSpec> {
	for (const useThousands of [false, true]) {
		for (const n of NEGATIVE_NUMBER_STYLES) {
			yield { category: 'number', decimals, useThousands, negativeStyle: n.style };
		}
	}
}

function* currencySpecs(decimals: number): Generator<NumberFormatSpec> {
	for (const symbol of CURRENCY_SYMBOL_PRESETS) {
		for (const n of NEGATIVE_NUMBER_STYLES) {
			yield { category: 'currency', decimals, currencySymbol: symbol.id, negativeStyle: n.style };
		}
	}
}

function* accountingSpecs(decimals: number): Generator<NumberFormatSpec> {
	for (const symbol of CURRENCY_SYMBOL_PRESETS) {
		yield { category: 'accounting', decimals, currencySymbol: symbol.id };
	}
}

function presetCode(code: string, presets: ReadonlyArray<{ code: string }>): string | undefined {
	return presets.find((p) => sameFormatCode(p.code, code))?.code;
}

function custom(code: string): NumberFormatChoice {
	return { category: 'custom', code };
}

/**
 * The Format Cells Number tab selection for `code`, like Excel: a code
 * that one of the standard categories produces selects that category with
 * its options; anything else shows under Custom.
 */
export function describeNumberFormat(code: string): NumberFormatChoice {
	const category = classifyFormat(code);
	switch (category) {
		case 'general':
			return { category };
		case 'text':
			return { category, code: '@' };
		case 'number':
		case 'currency':
		case 'accounting': {
			const decimals = clampDecimals(code);
			const specs =
				category === 'number'
					? numberSpecs(decimals)
					: category === 'currency'
						? currencySpecs(decimals)
						: accountingSpecs(decimals);
			const spec = firstMatch(code, specs);
			// The comma style is the accounting format without a symbol.
			return spec ? { ...spec, category } : custom(code);
		}
		case 'percentage':
		case 'scientific': {
			const spec = firstMatch(code, [{ category, decimals: clampDecimals(code) }]);
			return spec ? { ...spec, category } : custom(code);
		}
		case 'fraction': {
			const hit = presetCode(code, FRACTION_FORMAT_PRESETS);
			return hit ? { category, code: hit } : custom(code);
		}
		case 'date':
		case 'time': {
			const hit = presetCode(code, category === 'date' ? DATE_FORMAT_PRESETS : TIME_FORMAT_PRESETS);
			return hit ? { category, code: hit } : custom(code);
		}
		case 'special': {
			const hit = presetCode(code, SPECIAL_FORMAT_PRESETS);
			return hit ? { category, code: hit } : custom(code);
		}
		default:
			return custom(code);
	}
}

/** Built-in format ids in the order of Excel's Custom list (en-US). */
const CUSTOM_LIST_IDS = [
	0, 1, 2, 3, 4, 37, 38, 39, 40, 5, 6, 7, 8, 9, 10, 11, 48, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
	22, 45, 47, 49, 46, 42, 41, 44, 43,
];

/**
 * Codes listed by the Custom category: Excel's built-in formats followed by
 * the workbook's own custom formats (by id), without duplicates.
 */
export function customNumberFormatList(styles: Pick<StyleSheet, 'numFmts'>): string[] {
	const out: string[] = [];
	const seen = new Set<string>();
	const add = (code: string | undefined) => {
		if (code !== undefined && code !== '' && !seen.has(code)) {
			seen.add(code);
			out.push(code);
		}
	};
	for (const id of CUSTOM_LIST_IDS) {
		add(BUILTIN_NUMBER_FORMATS[id]);
	}
	const ids = Object.keys(styles.numFmts)
		.map(Number)
		.filter((id) => Number.isInteger(id))
		.sort((a, b) => a - b);
	for (const id of ids) {
		add(styles.numFmts[id]);
	}
	return out;
}

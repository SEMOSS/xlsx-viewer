/**
 * Number format engine: renders cell values through Excel format codes
 * (en-US output). Sections and conditions, colors, literals, `_x` spacing
 * and `*x` fills, digit placeholders, grouping and scaling, percent,
 * scientific, fractions, dates and times (elapsed, AM/PM, fractional
 * seconds, 1904 system), General and text sections. Format codes are parsed
 * once and cached; formatting a value is a short pass over a compiled
 * program.
 *
 * Exports are listed explicitly to keep the core barrel free of clashes.
 */
export { formatValue, formatValueForEdit } from './format-value';
export { formatGeneral, GENERAL_WIDTH as GENERAL_FORMAT_WIDTH } from './general';
export { parseFormatCode, clearFormatCache as clearNumberFormatCache } from './parse';
export { tokenizeFormat as tokenizeNumberFormat } from './tokenizer';
export { classifyFormat, getFormatDecimals, isDateFormat, isTimeOnlyFormat } from './classify';
export { adjustDecimals, MAX_FORMAT_DECIMALS } from './adjust-decimals';
export { buildNumberFormat } from './builder';
export type { NumberFormatSpec } from './builder';
export {
	CURRENCY_SYMBOL_PRESETS,
	DATE_FORMAT_PRESETS,
	FRACTION_FORMAT_PRESETS,
	NEGATIVE_NUMBER_STYLES,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from './presets';
export type {
	CurrencySymbolPreset,
	FormatPreset as NumberFormatPreset,
	LabeledFormatPreset as LabeledNumberFormatPreset,
	NegativeNumberStyle,
	NegativeNumberStyleOption,
} from './presets';
export type {
	FormatCategory,
	FormatOptions,
	FormattedValue,
	NumberFormatCondition,
	NumberFormatSection,
	NumberFormatToken,
	ParsedNumberFormat,
} from './types';

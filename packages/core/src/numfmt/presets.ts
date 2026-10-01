/**
 * Preset format codes offered by Excel's (en-US) Format Cells dialog.
 * `example` is the sample text the dialog lists for the preset.
 */

export interface FormatPreset {
	readonly code: string;
	readonly example: string;
}

export interface LabeledFormatPreset {
	readonly code: string;
	readonly label: string;
}

/** Date category (sample: Wednesday, March 14, 2012 1:30:55 PM). */
export const DATE_FORMAT_PRESETS: readonly FormatPreset[] = [
	{ code: 'm/d/yyyy', example: '3/14/2012' },
	{ code: '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy', example: 'Wednesday, March 14, 2012' },
	{ code: 'm/d;@', example: '3/14' },
	{ code: 'm/d/yy;@', example: '3/14/12' },
	{ code: 'mm/dd/yy;@', example: '03/14/12' },
	{ code: '[$-409]d-mmm;@', example: '14-Mar' },
	{ code: '[$-409]d-mmm-yy;@', example: '14-Mar-12' },
	{ code: '[$-409]dd-mmm-yy;@', example: '14-Mar-12' },
	{ code: '[$-409]mmm-yy;@', example: 'Mar-12' },
	{ code: '[$-409]mmmm-yy;@', example: 'March-12' },
	{ code: '[$-409]mmmm d, yyyy;@', example: 'March 14, 2012' },
	{ code: '[$-409]m/d/yy h:mm AM/PM;@', example: '3/14/12 1:30 PM' },
	{ code: 'm/d/yy h:mm;@', example: '3/14/12 13:30' },
	{ code: '[$-409]mmmmm;@', example: 'M' },
	{ code: '[$-409]mmmmm-yy;@', example: 'M-12' },
	{ code: 'm/d/yyyy;@', example: '3/14/2012' },
	{ code: '[$-409]d-mmm-yyyy;@', example: '14-Mar-2012' },
];

/** Time category (sample: 1:30:55.2 PM; the elapsed sample is 37:30:55). */
export const TIME_FORMAT_PRESETS: readonly FormatPreset[] = [
	{ code: '[$-F400]h:mm:ss\\ AM/PM', example: '1:30:55 PM' },
	{ code: 'h:mm;@', example: '13:30' },
	{ code: '[$-409]h:mm AM/PM;@', example: '1:30 PM' },
	{ code: 'h:mm:ss;@', example: '13:30:55' },
	{ code: '[$-409]h:mm:ss AM/PM;@', example: '1:30:55 PM' },
	{ code: 'mm:ss.0;@', example: '30:55.2' },
	{ code: '[h]:mm:ss;@', example: '37:30:55' },
	{ code: '[$-409]m/d/yy h:mm AM/PM;@', example: '3/14/12 1:30 PM' },
	{ code: 'm/d/yy h:mm;@', example: '3/14/12 13:30' },
];

/** Fraction category. */
export const FRACTION_FORMAT_PRESETS: readonly LabeledFormatPreset[] = [
	{ code: '# ?/?', label: 'Up to one digit (1/4)' },
	{ code: '# ??/??', label: 'Up to two digits (21/25)' },
	{ code: '# ???/???', label: 'Up to three digits (312/943)' },
	{ code: '# ?/2', label: 'As halves (1/2)' },
	{ code: '# ?/4', label: 'As quarters (2/4)' },
	{ code: '# ?/8', label: 'As eighths (4/8)' },
	{ code: '# ??/16', label: 'As sixteenths (8/16)' },
	{ code: '# ?/10', label: 'As tenths (3/10)' },
	{ code: '# ??/100', label: 'As hundredths (30/100)' },
];

/** Special category (en-US). */
export const SPECIAL_FORMAT_PRESETS: readonly LabeledFormatPreset[] = [
	{ code: '00000', label: 'Zip Code' },
	{ code: '00000-0000', label: 'Zip Code + 4' },
	{ code: '[<=9999999]###-####;(###) ###-####', label: 'Phone Number' },
	{ code: '000-00-0000', label: 'Social Security Number' },
];

/** A currency symbol choice: the code placed before and after the number. */
export interface CurrencySymbolPreset {
	readonly id: string;
	readonly label: string;
	readonly prefix: string;
	readonly suffix: string;
}

export const CURRENCY_SYMBOL_PRESETS: readonly CurrencySymbolPreset[] = [
	{ id: 'none', label: 'None', prefix: '', suffix: '' },
	{ id: '$', label: '$', prefix: '"$"', suffix: '' },
	{ id: '€', label: '€ Euro (€ 123)', prefix: '[$€-2]\\ ', suffix: '' },
	{ id: '€-after', label: '€ Euro (123 €)', prefix: '', suffix: '\\ [$€-1]' },
	{ id: '£', label: '£ English (United Kingdom)', prefix: '[$£-809]', suffix: '' },
	{ id: '¥', label: '¥ Japanese', prefix: '[$¥-411]', suffix: '' },
	{ id: '¥-CN', label: '¥ Chinese (PRC)', prefix: '[$¥-804]', suffix: '' },
	{ id: '₹', label: '₹ English (India)', prefix: '[$₹-4009]', suffix: '' },
	{ id: 'CHF', label: 'CHF German (Switzerland)', prefix: '[$CHF-807]\\ ', suffix: '' },
];

/** Negative number display choices of the Number and Currency categories. */
export type NegativeNumberStyle = 'minus' | 'red' | 'parens' | 'redParens' | 'redMinus';

export interface NegativeNumberStyleOption {
	readonly style: NegativeNumberStyle;
	readonly example: string;
}

export const NEGATIVE_NUMBER_STYLES: readonly NegativeNumberStyleOption[] = [
	{ style: 'minus', example: '-1234.10' },
	{ style: 'red', example: '1234.10' },
	{ style: 'parens', example: '(1234.10)' },
	{ style: 'redParens', example: '(1234.10)' },
	{ style: 'redMinus', example: '-1234.10' },
];

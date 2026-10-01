import type { FormatCategory, NegativeNumberStyle } from 'xlsx-viewer-core';
import {
	buildNumberFormat,
	DATE_FORMAT_PRESETS,
	describeNumberFormat,
	FRACTION_FORMAT_PRESETS,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from 'xlsx-viewer-core';

import type { TranslationKey } from '../../../../i18n';

/** Categories whose options are a list of preset types. */
export type TypeListCategory = 'date' | 'time' | 'fraction' | 'special';

/** Number tab state (all categories keep their own options while the dialog is open). */
export interface NumberDraft {
	category: FormatCategory;
	decimals: number;
	useThousands: boolean;
	negativeStyle: NegativeNumberStyle;
	currencySymbol: string;
	types: Record<TypeListCategory, string>;
	customCode: string;
}

/** The Category list, in Excel's order. */
export const NUMBER_CATEGORIES: ReadonlyArray<{ id: FormatCategory; labelKey: TranslationKey }> = [
	{ id: 'general', labelKey: 'dialog.formatCells.categoryName.general' },
	{ id: 'number', labelKey: 'dialog.formatCells.categoryName.number' },
	{ id: 'currency', labelKey: 'dialog.formatCells.categoryName.currency' },
	{ id: 'accounting', labelKey: 'dialog.formatCells.categoryName.accounting' },
	{ id: 'date', labelKey: 'dialog.formatCells.categoryName.date' },
	{ id: 'time', labelKey: 'dialog.formatCells.categoryName.time' },
	{ id: 'percentage', labelKey: 'dialog.formatCells.categoryName.percentage' },
	{ id: 'fraction', labelKey: 'dialog.formatCells.categoryName.fraction' },
	{ id: 'scientific', labelKey: 'dialog.formatCells.categoryName.scientific' },
	{ id: 'text', labelKey: 'dialog.formatCells.categoryName.text' },
	{ id: 'special', labelKey: 'dialog.formatCells.categoryName.special' },
	{ id: 'custom', labelKey: 'dialog.formatCells.categoryName.custom' },
];

/** Translation key of Excel's explanation under each category. */
export const CATEGORY_DESCRIPTIONS: Partial<Record<FormatCategory, TranslationKey>> = {
	general: 'dialog.formatCells.categoryDescription.general',
	number: 'dialog.formatCells.categoryDescription.number',
	currency: 'dialog.formatCells.categoryDescription.currency',
	accounting: 'dialog.formatCells.categoryDescription.accounting',
	date: 'dialog.formatCells.categoryDescription.date',
	time: 'dialog.formatCells.categoryDescription.time',
	percentage: 'dialog.formatCells.categoryDescription.percentage',
	text: 'dialog.formatCells.categoryDescription.text',
	special: 'dialog.formatCells.categoryDescription.special',
	custom: 'dialog.formatCells.categoryDescription.custom',
};

/**
 * Translation keys of the core's labelled presets (Fraction and Special types),
 * by format code. Codes without a key keep the core's English label.
 */
export const PRESET_LABEL_KEYS: Readonly<Partial<Record<string, TranslationKey>>> = {
	'# ?/?': 'dialog.formatCells.fractionType.oneDigit',
	'# ??/??': 'dialog.formatCells.fractionType.twoDigits',
	'# ???/???': 'dialog.formatCells.fractionType.threeDigits',
	'# ?/2': 'dialog.formatCells.fractionType.halves',
	'# ?/4': 'dialog.formatCells.fractionType.quarters',
	'# ?/8': 'dialog.formatCells.fractionType.eighths',
	'# ??/16': 'dialog.formatCells.fractionType.sixteenths',
	'# ?/10': 'dialog.formatCells.fractionType.tenths',
	'# ??/100': 'dialog.formatCells.fractionType.hundredths',
	'00000': 'dialog.formatCells.specialType.zipCode',
	'00000-0000': 'dialog.formatCells.specialType.zipCodePlus4',
	'[<=9999999]###-####;(###) ###-####': 'dialog.formatCells.specialType.phoneNumber',
	'000-00-0000': 'dialog.formatCells.specialType.socialSecurityNumber',
};

/** Translation keys of the currency symbol choices, by preset id (`$` needs none). */
export const CURRENCY_LABEL_KEYS: Readonly<Partial<Record<string, TranslationKey>>> = {
	none: 'dialog.none',
	'€': 'dialog.formatCells.currencySymbol.euroPrefix',
	'€-after': 'dialog.formatCells.currencySymbol.euroSuffix',
	'£': 'dialog.formatCells.currencySymbol.pound',
	'¥': 'dialog.formatCells.currencySymbol.yen',
	'¥-CN': 'dialog.formatCells.currencySymbol.yuan',
	'₹': 'dialog.formatCells.currencySymbol.rupee',
	CHF: 'dialog.formatCells.currencySymbol.swissFranc',
};

/** Codes that follow the operating system's regional settings (listed with an asterisk). */
export const SYSTEM_FORMAT_CODES: ReadonlySet<string> = new Set([
	DATE_FORMAT_PRESETS[0].code,
	DATE_FORMAT_PRESETS[1].code,
	TIME_FORMAT_PRESETS[0].code,
]);

const DEFAULT_TYPES: Record<TypeListCategory, string> = {
	date: DATE_FORMAT_PRESETS[0].code,
	time: TIME_FORMAT_PRESETS[0].code,
	fraction: FRACTION_FORMAT_PRESETS[0].code,
	special: SPECIAL_FORMAT_PRESETS[0].code,
};

export function isTypeListCategory(c: FormatCategory): c is TypeListCategory {
	return c === 'date' || c === 'time' || c === 'fraction' || c === 'special';
}

/** The draft the dialog opens with for a format code. */
export function numberDraftFrom(code: string): NumberDraft {
	const choice = describeNumberFormat(code);
	const draft: NumberDraft = {
		category: choice.category,
		decimals: choice.decimals ?? 2,
		useThousands: choice.useThousands ?? false,
		negativeStyle: choice.negativeStyle ?? 'minus',
		currencySymbol: choice.currencySymbol ?? '$',
		types: { ...DEFAULT_TYPES },
		customCode: code,
	};
	if (isTypeListCategory(choice.category) && choice.code) {
		draft.types[choice.category] = choice.code;
	}
	return draft;
}

/** Format code the draft describes. */
export function numberDraftCode(d: NumberDraft): string {
	switch (d.category) {
		case 'general':
			return 'General';
		case 'text':
			return '@';
		case 'custom':
			return d.customCode.trim() === '' ? 'General' : d.customCode;
		case 'date':
		case 'time':
		case 'fraction':
		case 'special':
			return d.types[d.category];
		default:
			return buildNumberFormat({
				category: d.category,
				decimals: d.decimals,
				useThousands: d.useThousands,
				negativeStyle: d.negativeStyle,
				currencySymbol: d.currencySymbol,
			});
	}
}

/** Switch category; entering Custom starts from the code shown so far, like Excel. */
export function withCategory(d: NumberDraft, category: FormatCategory): NumberDraft {
	if (category === d.category) {
		return d;
	}
	const next = { ...d, category };
	if (category === 'custom') {
		next.customCode = numberDraftCode(d);
	}
	return next;
}

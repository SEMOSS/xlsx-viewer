import type {
	CfRule,
	CfTimePeriod,
	HighlightPresetName,
	RangeAddress,
	Worksheet,
} from 'xlsx-viewer-core';
import {
	cellKey,
	createAverageRule,
	createCellIsRule,
	createDateOccurringRule,
	createDuplicateRule,
	createTextRule,
	createTopBottomRule,
	keyCol,
	keyRow,
	numberToText,
	toCriteriaFormula,
} from 'xlsx-viewer-core';

import { createTranslator, type Translate, type TranslationKey } from '../../i18n';

/** The rule dialogs of Home > Conditional Formatting > Highlight Cells / Top Bottom Rules. */
export type ConditionalRuleKind =
	| 'greaterThan'
	| 'lessThan'
	| 'between'
	| 'equalTo'
	| 'textContains'
	| 'dateOccurring'
	| 'duplicateValues'
	| 'top10'
	| 'topPercent'
	| 'bottom10'
	| 'bottomPercent'
	| 'aboveAverage'
	| 'belowAverage';

export interface ConditionalRuleSpec {
	/** The dialog title. */
	titleKey: TranslationKey;
	/** "Format cells that ...", above the inputs (also their accessible name). */
	promptKey: TranslationKey;
	input: 'value' | 'between' | 'text' | 'date' | 'duplicate' | 'rank' | 'none';
	/** Rank is a percentage (Top / Bottom 10%). */
	percent?: boolean;
}

export const CONDITIONAL_RULES: Readonly<Record<ConditionalRuleKind, ConditionalRuleSpec>> = {
	greaterThan: {
		titleKey: 'dialog.cfRule.title.greaterThan',
		promptKey: 'dialog.cfRule.prompt.greaterThan',
		input: 'value',
	},
	lessThan: {
		titleKey: 'dialog.cfRule.title.lessThan',
		promptKey: 'dialog.cfRule.prompt.lessThan',
		input: 'value',
	},
	between: {
		titleKey: 'dialog.cfRule.title.between',
		promptKey: 'dialog.cfRule.prompt.between',
		input: 'between',
	},
	equalTo: {
		titleKey: 'dialog.cfRule.title.equalTo',
		promptKey: 'dialog.cfRule.prompt.equalTo',
		input: 'value',
	},
	textContains: {
		titleKey: 'dialog.cfRule.title.textContains',
		promptKey: 'dialog.cfRule.prompt.textContains',
		input: 'text',
	},
	dateOccurring: {
		titleKey: 'dialog.cfRule.title.dateOccurring',
		promptKey: 'dialog.cfRule.prompt.dateOccurring',
		input: 'date',
	},
	duplicateValues: {
		titleKey: 'dialog.cfRule.title.duplicateValues',
		promptKey: 'dialog.cfRule.prompt.duplicateValues',
		input: 'duplicate',
	},
	top10: {
		titleKey: 'dialog.cfRule.title.top10',
		promptKey: 'dialog.cfRule.prompt.top',
		input: 'rank',
	},
	topPercent: {
		titleKey: 'dialog.cfRule.title.topPercent',
		promptKey: 'dialog.cfRule.prompt.top',
		input: 'rank',
		percent: true,
	},
	bottom10: {
		titleKey: 'dialog.cfRule.title.bottom10',
		promptKey: 'dialog.cfRule.prompt.bottom',
		input: 'rank',
	},
	bottomPercent: {
		titleKey: 'dialog.cfRule.title.bottomPercent',
		promptKey: 'dialog.cfRule.prompt.bottom',
		input: 'rank',
		percent: true,
	},
	aboveAverage: {
		titleKey: 'dialog.cfRule.title.aboveAverage',
		promptKey: 'dialog.cfRule.prompt.aboveAverage',
		input: 'none',
	},
	belowAverage: {
		titleKey: 'dialog.cfRule.title.belowAverage',
		promptKey: 'dialog.cfRule.prompt.belowAverage',
		input: 'none',
	},
};

export function isConditionalRuleKind(v: unknown): v is ConditionalRuleKind {
	return typeof v === 'string' && v in CONDITIONAL_RULES;
}

/** "A Date Occurring" choices, in Excel's order. */
export const DATE_PERIODS: ReadonlyArray<{ value: CfTimePeriod; labelKey: TranslationKey }> = [
	{ value: 'yesterday', labelKey: 'dialog.cfRule.period.yesterday' },
	{ value: 'today', labelKey: 'dialog.cfRule.period.today' },
	{ value: 'tomorrow', labelKey: 'dialog.cfRule.period.tomorrow' },
	{ value: 'last7Days', labelKey: 'dialog.cfRule.period.last7Days' },
	{ value: 'lastWeek', labelKey: 'dialog.cfRule.period.lastWeek' },
	{ value: 'thisWeek', labelKey: 'dialog.cfRule.period.thisWeek' },
	{ value: 'nextWeek', labelKey: 'dialog.cfRule.period.nextWeek' },
	{ value: 'lastMonth', labelKey: 'dialog.cfRule.period.lastMonth' },
	{ value: 'thisMonth', labelKey: 'dialog.cfRule.period.thisMonth' },
	{ value: 'nextMonth', labelKey: 'dialog.cfRule.period.nextMonth' },
];

/** Captions of core's highlight presets (the "with" box), in the UI language. */
export const HIGHLIGHT_PRESET_KEYS: Readonly<Record<HighlightPresetName, TranslationKey>> = {
	lightRedFillDarkRedText: 'dialog.cfRule.preset.lightRedFillDarkRedText',
	yellowFillDarkYellowText: 'dialog.cfRule.preset.yellowFillDarkYellowText',
	greenFillDarkGreenText: 'dialog.cfRule.preset.greenFillDarkGreenText',
	lightRedFill: 'dialog.cfRule.preset.lightRedFill',
	redText: 'dialog.cfRule.preset.redText',
	redBorder: 'dialog.cfRule.preset.redBorder',
};

export interface ConditionalRuleInput {
	/** Value boxes (the second one only for Between). */
	values: [string, string];
	text: string;
	period: CfTimePeriod;
	/** Duplicate Values: highlight unique values instead. */
	unique: boolean;
	rank: string;
	format: HighlightPresetName;
}

/** Build the rule a dialog describes, or Excel's complaint about the input (worded by `t`). */
export function buildConditionalRule(
	kind: ConditionalRuleKind,
	input: ConditionalRuleInput,
	date1904 = false,
	t: Translate = createTranslator(),
): { rule: CfRule } | { error: string } {
	const spec = CONDITIONAL_RULES[kind];
	const criteria = (text: string) => toCriteriaFormula(text, date1904);
	switch (spec.input) {
		case 'value':
		case 'between': {
			const count = spec.input === 'between' ? 2 : 1;
			const values = input.values.slice(0, count).map((v) => v.trim());
			if (values.some((v) => v === '' || v === '=')) {
				return { error: t('dialog.cfRule.enterValue') };
			}
			const operator =
				kind === 'greaterThan'
					? 'greaterThan'
					: kind === 'lessThan'
						? 'lessThan'
						: kind === 'between'
							? 'between'
							: 'equal';
			return { rule: createCellIsRule(operator, values.map(criteria), input.format) };
		}
		case 'text':
			if (input.text === '') {
				return { error: t('dialog.cfRule.enterText') };
			}
			return { rule: createTextRule('containsText', input.text, input.format) };
		case 'date':
			return { rule: createDateOccurringRule(input.period, input.format) };
		case 'duplicate':
			return { rule: createDuplicateRule(input.unique, input.format) };
		case 'rank': {
			const n = Number(input.rank.trim());
			const max = spec.percent ? 100 : 1000;
			if (!Number.isInteger(n) || n < 1 || n > max) {
				return { error: t('dialog.cfRule.rankRange', { min: 1, max }) };
			}
			return {
				rule: createTopBottomRule(
					{ rank: n, percent: spec.percent, bottom: kind.startsWith('bottom') },
					input.format,
				),
			};
		}
		default:
			return { rule: createAverageRule(kind === 'aboveAverage', input.format) };
	}
}

/** Numbers in the ranges (stored cells only), up to `limit`. */
function numbersIn(sheet: Worksheet, ranges: readonly RangeAddress[], limit = 20000): number[] {
	const out: number[] = [];
	for (const [key, cell] of sheet.cells) {
		if (typeof cell.value !== 'number') {
			continue;
		}
		const row = keyRow(key);
		const col = keyCol(key);
		if (ranges.some((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)) {
			out.push(cell.value);
			if (out.length >= limit) {
				break;
			}
		}
	}
	return out.sort((a, b) => a - b);
}

function rounded(n: number): string {
	return numberToText(Math.round(n * 100) / 100);
}

function quantile(sorted: readonly number[], q: number): number {
	const pos = (sorted.length - 1) * q;
	const lo = Math.floor(pos);
	const hi = Math.ceil(pos);
	return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/**
 * Values a dialog opens with, taken from the selected data like Excel
 * pre-fills them: the average for Greater Than / Less Than / Equal To,
 * the lower and upper quartiles for Between, the active cell's text for
 * Text that Contains.
 */
export function suggestRuleValues(
	kind: ConditionalRuleKind,
	sheet: Worksheet,
	ranges: readonly RangeAddress[],
	active: { row: number; col: number },
): { values: [string, string]; text: string } {
	const activeValue = sheet.cells.get(cellKey(active.row, active.col))?.value;
	const text = typeof activeValue === 'string' ? activeValue : '';
	const spec = CONDITIONAL_RULES[kind];
	if (spec.input !== 'value' && spec.input !== 'between') {
		return { values: ['', ''], text };
	}
	const nums = numbersIn(sheet, ranges);
	if (nums.length === 0) {
		return { values: ['', ''], text };
	}
	if (spec.input === 'between') {
		return { values: [rounded(quantile(nums, 0.25)), rounded(quantile(nums, 0.75))], text };
	}
	const avg = nums.reduce((s, n) => s + n, 0) / nums.length;
	return { values: [rounded(avg), ''], text };
}

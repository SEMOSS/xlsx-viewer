import { numberToText } from '../formula/coerce';
import type {
	CfOperator,
	CfRule,
	CfTimePeriod,
	Cfvo,
	IconSetName,
} from '../types/conditional-format';
import type { ColorSpec, DifferentialFormat } from '../types/style';
import { parseNumericText } from '../values/parse-number';
import { stripLeadingEquals } from './formula-eval';
import { defaultIconCfvos } from './icon-set';
import { ICON_SETS } from './icon-sets';
import { COLOR_SCALE_PRESETS, HIGHLIGHT_PRESETS, type HighlightPresetName } from './presets';
import type { TextRuleKind } from './text-match';

/** A preset name or an explicit differential format. */
export type HighlightFormat = HighlightPresetName | DifferentialFormat;

function cloneDxf(dxf: DifferentialFormat): DifferentialFormat {
	return JSON.parse(JSON.stringify(dxf)) as DifferentialFormat;
}

function resolveFormat(format: HighlightFormat): DifferentialFormat {
	return cloneDxf(typeof format === 'string' ? HIGHLIGHT_PRESETS[format] : format);
}

/**
 * Turn dialog input into a criteria formula the way Excel does: `=...` is a
 * formula, numbers / dates / percentages / booleans are written as values and
 * anything else becomes a quoted string literal.
 */
export function toCriteriaFormula(input: string | number | boolean, date1904 = false): string {
	if (typeof input === 'number') {
		return numberToText(input);
	}
	if (typeof input === 'boolean') {
		return input ? 'TRUE' : 'FALSE';
	}
	const t = input.trim();
	if (t.startsWith('=')) {
		return stripLeadingEquals(t);
	}
	const parsed = parseNumericText(t, { date1904 });
	if (parsed) {
		return numberToText(parsed.value);
	}
	if (/^(TRUE|FALSE)$/i.test(t)) {
		return t.toUpperCase();
	}
	return `"${input.replace(/"/g, '""')}"`;
}

/**
 * "Highlight Cells Rules": Greater Than, Less Than, Between, Equal To, ...
 * `formulas` are criteria formulas (see {@link toCriteriaFormula}).
 * New rules get priority 1; {@link addConditionalFormat} renumbers.
 */
export function createCellIsRule(
	operator: CfOperator,
	formulas: readonly string[],
	format: HighlightFormat,
): CfRule {
	return {
		type: 'cellIs',
		priority: 1,
		operator,
		formulas: formulas.map(stripLeadingEquals),
		dxf: resolveFormat(format),
	};
}

const TEXT_RULE_OPERATORS: Record<TextRuleKind, CfOperator> = {
	containsText: 'containsText',
	notContainsText: 'notContains',
	beginsWith: 'beginsWith',
	endsWith: 'endsWith',
};

/**
 * "Text that Contains" (and does not contain / begins with / ends with).
 * The formula Excel stores with it is added by `addConditionalFormat`.
 */
export function createTextRule(kind: TextRuleKind, text: string, format: HighlightFormat): CfRule {
	return {
		type: kind,
		priority: 1,
		operator: TEXT_RULE_OPERATORS[kind],
		text,
		dxf: resolveFormat(format),
	};
}

/** "A Date Occurring". */
export function createDateOccurringRule(period: CfTimePeriod, format: HighlightFormat): CfRule {
	return { type: 'timePeriod', priority: 1, timePeriod: period, dxf: resolveFormat(format) };
}

/** "Duplicate Values" (or unique values when `unique`). */
export function createDuplicateRule(unique: boolean, format: HighlightFormat): CfRule {
	return {
		type: unique ? 'uniqueValues' : 'duplicateValues',
		priority: 1,
		dxf: resolveFormat(format),
	};
}

export interface TopBottomOptions {
	/** Bottom items instead of top. Default false. */
	bottom?: boolean;
	/** `rank` is a percentage of the numeric cells. Default false. */
	percent?: boolean;
	/** Number (or percentage) of items. Default 10. */
	rank?: number;
}

/** "Top 10 Items", "Top 10%", "Bottom 10 Items", "Bottom 10%". */
export function createTopBottomRule(options: TopBottomOptions, format: HighlightFormat): CfRule {
	const rule: CfRule = {
		type: 'top10',
		priority: 1,
		rank: Math.max(1, Math.floor(options.rank ?? 10)),
		dxf: resolveFormat(format),
	};
	if (options.percent) {
		rule.percent = true;
	}
	if (options.bottom) {
		rule.bottom = true;
	}
	return rule;
}

export interface AverageRuleOptions {
	/** Include values equal to the average ("equal or above"). */
	equalAverage?: boolean;
	/** Number of standard deviations above / below the average (1-3 in Excel's dialog). */
	stdDev?: number;
}

/** "Above Average" / "Below Average" (optionally "equal or ..." or N std dev). */
export function createAverageRule(
	above: boolean,
	format: HighlightFormat,
	options: AverageRuleOptions = {},
): CfRule {
	const rule: CfRule = { type: 'aboveAverage', priority: 1, dxf: resolveFormat(format) };
	if (!above) {
		rule.aboveAverage = false;
	}
	if (options.equalAverage) {
		rule.equalAverage = true;
	}
	if (options.stdDev && options.stdDev > 0) {
		rule.stdDev = options.stdDev;
	}
	return rule;
}

function toColorSpec(color: string | ColorSpec): ColorSpec {
	if (typeof color !== 'string') {
		return { ...color };
	}
	const hex = color.replace(/^#/, '').toUpperCase();
	return { rgb: hex.length === 6 ? `FF${hex}` : hex };
}

/**
 * Data bar as Excel 2010+ creates it from the gallery: automatic minimum and
 * maximum, full-width scale, negative values in red on an automatic axis.
 * Gradient bars get a border in the bar color, solid bars none.
 */
export function createDataBarRule(color: string | ColorSpec, gradient = true): CfRule {
	const spec = toColorSpec(color);
	const red: ColorSpec = { rgb: 'FFFF0000' };
	return {
		type: 'dataBar',
		priority: 1,
		dataBar: {
			cfvos: [{ type: 'autoMin' }, { type: 'autoMax' }],
			color: spec,
			minLength: 0,
			maxLength: 100,
			gradient,
			border: gradient,
			...(gradient ? { borderColor: { ...spec }, negativeBorderColor: { ...red } } : {}),
			negativeFillColor: red,
			axisPosition: 'automatic',
			axisColor: { rgb: 'FF000000' },
			direction: 'context',
		},
	};
}

/** One of Excel's twelve color scale presets (index into `COLOR_SCALE_PRESETS`). */
export function createColorScaleRule(presetIndex: number): CfRule {
	const preset = COLOR_SCALE_PRESETS[presetIndex];
	if (!preset) {
		throw new RangeError(`createColorScaleRule: unknown preset ${presetIndex}`);
	}
	const cfvos: Cfvo[] =
		preset.colors.length === 3
			? [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }]
			: [{ type: 'min' }, { type: 'max' }];
	return {
		type: 'colorScale',
		priority: 1,
		colorScale: { cfvos, colors: preset.colors.map((rgb) => ({ rgb })) },
	};
}

/** Icon set with Excel's default percent thresholds. */
export function createIconSetRule(
	iconSet: IconSetName,
	options: { reverse?: boolean; showValue?: boolean } = {},
): CfRule {
	const count = ICON_SETS[iconSet]?.count ?? 0;
	if (count < 3) {
		throw new RangeError(`createIconSetRule: ${iconSet} is not a usable icon set`);
	}
	const rule: CfRule = {
		type: 'iconSet',
		priority: 1,
		iconSet: { iconSet, cfvos: defaultIconCfvos(count) },
	};
	if (options.reverse) {
		rule.iconSet!.reverse = true;
	}
	if (options.showValue === false) {
		rule.iconSet!.showValue = false;
	}
	return rule;
}

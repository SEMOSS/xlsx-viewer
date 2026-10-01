import type { RangeAddress } from './address';
import type { ColorSpec, DifferentialFormat } from './style';

export type CfRuleType =
	| 'expression'
	| 'cellIs'
	| 'colorScale'
	| 'dataBar'
	| 'iconSet'
	| 'top10'
	| 'uniqueValues'
	| 'duplicateValues'
	| 'containsText'
	| 'notContainsText'
	| 'beginsWith'
	| 'endsWith'
	| 'containsBlanks'
	| 'notContainsBlanks'
	| 'containsErrors'
	| 'notContainsErrors'
	| 'timePeriod'
	| 'aboveAverage';

export type CfOperator =
	| 'lessThan'
	| 'lessThanOrEqual'
	| 'equal'
	| 'notEqual'
	| 'greaterThanOrEqual'
	| 'greaterThan'
	| 'between'
	| 'notBetween'
	| 'containsText'
	| 'notContains'
	| 'beginsWith'
	| 'endsWith';

export type CfTimePeriod =
	| 'today'
	| 'yesterday'
	| 'tomorrow'
	| 'last7Days'
	| 'thisMonth'
	| 'lastMonth'
	| 'nextMonth'
	| 'thisWeek'
	| 'lastWeek'
	| 'nextWeek';

/** Conditional formatting value object (threshold) for color scales, data bars and icon sets. */
export interface Cfvo {
	type: 'num' | 'percent' | 'max' | 'min' | 'formula' | 'percentile' | 'autoMin' | 'autoMax';
	/** Number or formula text (no leading `=`), depending on `type`. */
	value?: string;
	/** Icon sets: threshold comparison is `>=` (default) rather than `>`. */
	gte?: boolean;
}

export interface CfColorScale {
	cfvos: Cfvo[];
	/** One color per cfvo (2 or 3). */
	colors: ColorSpec[];
}

export interface CfDataBar {
	cfvos: [Cfvo, Cfvo];
	color: ColorSpec;
	showValue?: boolean;
	/** Minimum / maximum bar length as a percentage of the cell width. */
	minLength?: number;
	maxLength?: number;
	/** Excel 2010+ extension properties (from `x14:dataBar`). */
	gradient?: boolean;
	border?: boolean;
	borderColor?: ColorSpec;
	negativeFillColor?: ColorSpec;
	negativeBorderColor?: ColorSpec;
	negativeBarColorSameAsPositive?: boolean;
	axisPosition?: 'automatic' | 'middle' | 'none';
	axisColor?: ColorSpec;
	direction?: 'context' | 'leftToRight' | 'rightToLeft';
}

export type IconSetName =
	| '3Arrows'
	| '3ArrowsGray'
	| '3Flags'
	| '3TrafficLights1'
	| '3TrafficLights2'
	| '3Signs'
	| '3Symbols'
	| '3Symbols2'
	| '3Stars'
	| '3Triangles'
	| '4Arrows'
	| '4ArrowsGray'
	| '4RedToBlack'
	| '4Rating'
	| '4TrafficLights'
	| '5Arrows'
	| '5ArrowsGray'
	| '5Rating'
	| '5Quarters'
	| '5Boxes'
	| 'NoIcons';

export interface CfIconSet {
	iconSet: IconSetName;
	cfvos: Cfvo[];
	showValue?: boolean;
	reverse?: boolean;
	percent?: boolean;
	/** Per-threshold custom icons (Excel 2010+), as `[iconSet, iconIndex]`. */
	customIcons?: Array<{ iconSet: IconSetName; iconId: number }>;
}

export interface CfRule {
	type: CfRuleType;
	priority: number;
	stopIfTrue?: boolean;
	/** Index into `workbook.styles.dxfs`. */
	dxfId?: number;
	/** Inline differential format (Excel 2010+ rules stored in `extLst`). */
	dxf?: DifferentialFormat;
	operator?: CfOperator;
	/** Rule formulas in display form (no leading `=`), relative to the top-left cell of the first range. */
	formulas?: string[];
	text?: string;
	timePeriod?: CfTimePeriod;
	rank?: number;
	percent?: boolean;
	bottom?: boolean;
	aboveAverage?: boolean;
	equalAverage?: boolean;
	stdDev?: number;
	colorScale?: CfColorScale;
	dataBar?: CfDataBar;
	iconSet?: CfIconSet;
	/** `x14:id` linking a 2007 rule with its 2010 extension counterpart (round-trip). */
	extId?: string;
	/** The rule lives only in the Excel 2010+ `extLst` block. */
	fromExtension?: boolean;
}

export interface ConditionalFormat {
	ranges: RangeAddress[];
	rules: CfRule[];
	pivot?: boolean;
}

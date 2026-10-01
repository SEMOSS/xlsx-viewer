/**
 * Conditional formatting: evaluation for rendering (highlight rules, color
 * scales, data bars, icon sets), icon descriptors, ribbon presets / rule
 * factories and rule editing helpers.
 */
export type {
	CfCellResult,
	ConditionalFormatEngineOptions,
	DataBarResult,
	IconResult,
} from './types';
export type { FormulaEvaluator } from './formula-eval';
export { ConditionalFormatEngine } from './engine';
export {
	ICON_COLORS,
	ICON_SET_GROUPS,
	ICON_SETS,
	type IconDescriptor,
	type IconSetDescriptor,
	type IconShape,
} from './icon-sets';
export {
	COLOR_SCALE_PRESETS,
	DATA_BAR_PRESETS,
	HIGHLIGHT_PRESET_LABELS,
	HIGHLIGHT_PRESETS,
	type ColorScalePreset,
	type DataBarPreset,
	type HighlightPresetName,
} from './presets';
export {
	createAverageRule,
	createCellIsRule,
	createColorScaleRule,
	createDataBarRule,
	createDateOccurringRule,
	createDuplicateRule,
	createIconSetRule,
	createTextRule,
	createTopBottomRule,
	toCriteriaFormula,
	type AverageRuleOptions,
	type HighlightFormat,
	type TopBottomOptions,
} from './factories';
export type { TextRuleKind } from './text-match';
export { buildConditionalRuleFormula } from './rule-formulas';
export { addConditionalFormat, clearConditionalFormats, getConditionalFormatRules } from './edit';
export { dxfBackgroundColor, mergeDifferentialFormats } from './dxf';
export {
	layoutDataBar,
	type DataBarAxisMode,
	type DataBarLayout,
	type DataBarLayoutInput,
} from './data-bar';

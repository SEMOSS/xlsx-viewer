/**
 * Formula language: tokenizer, parser, serializer and reference rewriting
 * (copy translation, row / column insertion and deletion, moves, sheet
 * renames, R1C1) plus formula-bar helpers. No evaluation happens here.
 *
 * `translateFormula` and the prefix helpers are exported from
 * `formula/translate.ts` and `formula/prefixes.ts`.
 */
export type { FormulaToken, FormulaTokenType, RefArea, RefKind, TokenizeOptions } from './tokens';
export { tokenizeFormula } from './tokenizer';
export { FormulaParseError } from './parse-error';
export { MAX_FUNCTION_ARGS, parseFormula, parseFormulaTokens, type ParseOptions } from './parser';
export {
	formatFormulaNumber,
	formatFormulaString,
	serializeFormula,
	type SerializeOptions,
} from './serializer';
export {
	formatRefArea,
	formatRefCorner,
	formatReference,
	formatSheetQualifier,
	type ParsedReference,
	parseReferenceText,
	refAreaToRange,
	sameSheetName,
} from './references';
export {
	escapeStructuredColumn,
	formatStructuredReference,
	parseStructuredReference,
} from './structured';
export { normalizeFunctionName, stripFilePrefixes } from './names';
export {
	adjustAreaForStructureChange,
	adjustRangeForStructureChange,
	type FormulaAdjustment,
	type StructureChange,
} from './adjust-range';
export { adjustFormulaForStructureChange } from './adjust';
export {
	adjustFormulaForMove,
	type FormulaMoveOptions,
	type FormulaMoveSource,
	type FormulaMoveTarget,
} from './move';
export {
	formulaReferencesSheet,
	invalidateSheetInFormula,
	renameSheetInFormula,
} from './sheet-refs';
export {
	type FormulaDisplayToken,
	type FormulaDisplayTokenKind,
	type FormulaReference,
	getFormulaReferences,
	tokenizeForDisplay,
} from './editor-display';
export {
	type ActiveFunctionCall,
	type CompletionPrefix,
	getActiveFunctionCall,
	getCompletionPrefix,
	isPointModeAllowed,
	type ReferenceToggle,
	toggleReferenceAbsolute,
} from './editor-caret';
export { a1ToR1C1, formatRefAreaR1C1, r1c1ToA1 } from './r1c1';
export { analyzeCalls, type FormulaCallInfo } from './token-utils';

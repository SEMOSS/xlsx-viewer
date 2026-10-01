/**
 * Formula evaluator: walks a parsed formula (see `formula/language`) against
 * an {@link EvaluationHost}. Operators broadcast over arrays, references
 * stay lazy, names / tables / LET / LAMBDA are resolved here and function
 * calls go through `invokeFunction`.
 */
export { MAX_EVAL_DEPTH } from './calls';
export { Evaluator } from './evaluator';
export { EvalFrame, type FrameOwner } from './frame';
export type { EvaluationHost, HostTable, NameDefinition } from './host';
export {
	etaDefinition,
	etaLambda,
	isOmittedValue,
	OMITTED,
	optionalParamName,
	paramKey,
} from './lambda';
export {
	applyOperator,
	broadcastBinary,
	compareNumbers,
	mapArray,
	negate,
	numberResult,
	percent,
	type ValueOperator,
} from './operators';
export {
	cellNodeRange,
	intersectionOperator,
	intersectRef,
	makeRef,
	rangeNodeRange,
	rangeOperator,
	resolveSheetSpan,
	type SheetSpan,
	unionOperator,
} from './references';
export {
	arrayCellResult,
	cellArray,
	cellScalar,
	legacyCellResult,
	type ResultReader,
} from './results';
export { structuredRange, tableColumnIndex, tableDataRange, tableDataRows } from './structured';

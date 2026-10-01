/**
 * Calculation engine: dependency tracking and recalculation of formulas
 * (see {@link CalcEngine}).
 */
export { CalcEngine } from './engine';
export type {
	CalcEngineOptions,
	CellRef,
	EvaluateOptions,
	PrecedentArea,
	RecalcResult,
} from './types';

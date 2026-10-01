import type { RangeAddress } from '../types/address';

/** A cell position within the workbook. */
export interface CellRef {
	sheet: number;
	row: number;
	col: number;
}

export interface RecalcResult {
	/** Cells whose value changed (including spill ranges that grew or shrank). */
	changed: CellRef[];
	/** Cells that are part of a circular reference. */
	circular: CellRef[];
}

export interface CalcEngineOptions {
	/** Clock used by NOW / TODAY (testing aid). */
	now?: () => Date;
	/** Random source used by RAND / RANDBETWEEN / RANDARRAY (testing aid). */
	random?: () => number;
}

export interface EvaluateOptions {
	/** Evaluate with dynamic-array semantics (spill) instead of implicit intersection. Default true. */
	dynamicArrays?: boolean;
}

/** A block of cells a formula reads (for Trace Precedents). */
export interface PrecedentArea {
	sheet: number;
	range: RangeAddress;
}

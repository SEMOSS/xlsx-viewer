import type { RangeAddress } from '../types/address';
import type { ParsedFormula } from './formula-cache';
import type { RangeIndexEntry } from './range-index';

/** How a formula cell is evaluated. */
export type FormulaMode = 'legacy' | 'dynamic' | 'cse';

/** A distinct range read by one or more formulas (shared to keep the reverse index small). */
export interface RangeDep {
	readonly sheet: number;
	readonly range: RangeAddress;
	readonly id: string;
	readonly dependents: Set<FormulaNode>;
	entry: RangeIndexEntry<RangeDep> | null;
}

/** The precedents registered for a node (kept for unregistering). */
export interface NodeDeps {
	/** Flattened (sheet, cellKey) pairs. */
	readonly cells: number[];
	readonly ranges: RangeDep[];
}

/** One formula cell in the dependency graph, plus its recalculation state. */
export class FormulaNode {
	deps: NodeDeps = { cells: [], ranges: [] };
	volatile = false;
	/** Uses an unsupported function: the cached value is kept and never recalculated. */
	frozen = false;
	/** Contains SUBTOTAL / AGGREGATE. */
	subtotal = false;

	/** Must be recalculated in the current pass. */
	dirty = false;
	/** Recalculation counter of the last time the node was evaluated (oscillation guard). */
	recalcId = 0;
	evalCount = 0;
	/** On the evaluation retry stack. */
	inEval = false;

	// Tarjan bookkeeping.
	tStamp = 0;
	tIndex = 0;
	tLow = 0;
	onStack = false;

	/** Current multi-cell spill range of a dynamic-array formula. */
	spill: RangeAddress | null = null;
	spillEntry: RangeIndexEntry<FormulaNode> | null = null;
	/** Range a blocked dynamic-array formula (#SPILL!) wants to fill. */
	blocked: RangeAddress | null = null;
	blockedEntry: RangeIndexEntry<FormulaNode> | null = null;
	/** The last result was an array (so `A1#` is valid even for a 1x1 result). */
	arrayResult = false;
	/** Block of a legacy CSE array formula (indexed with the spills: its cells are outputs too). */
	cse: RangeAddress | null = null;
	cseEntry: RangeIndexEntry<FormulaNode> | null = null;

	constructor(
		readonly sheet: number,
		readonly row: number,
		readonly col: number,
		readonly key: number,
		readonly formula: string,
		readonly parsed: ParsedFormula,
		readonly mode: FormulaMode,
	) {}
}

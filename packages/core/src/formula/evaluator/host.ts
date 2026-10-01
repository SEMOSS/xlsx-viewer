import type { RangeAddress } from '../../types/address';
import type { CellError } from '../../types/cell';
import type { Table } from '../../types/table';
import type { AstNode } from '../ast';
import type { CellInfoProvider } from '../functions/information/cell';
import type { FunctionRegistry } from '../registry';
import type { ArrayValue, RefValue, Scalar } from '../types';

/** A defined name as seen by the evaluator. */
export interface NameDefinition {
	/** Name as written in the workbook. */
	readonly name: string;
	/** Parsed formula, or `null` when the formula could not be parsed. */
	readonly ast: AstNode | null;
	/** Sheet index of a sheet-scoped name; `undefined` for workbook scope. */
	readonly localSheet?: number;
}

/** A table together with the index of the sheet that holds it. */
export interface HostTable {
	readonly sheet: number;
	readonly table: Table;
}

/**
 * Everything the AST evaluator needs from its surroundings: cell access,
 * sheet, name and table lookups, spill ranges and the environment (clock,
 * random source). The calculation engine implements it over a workbook.
 *
 * Cell accessors may throw to suspend an evaluation (the engine does this
 * when a formula reads a cell that still has to be recalculated); the
 * evaluator never catches such exceptions.
 */
export interface EvaluationHost {
	readonly registry: FunctionRegistry;
	readonly date1904: boolean;

	getCellValue(sheet: number, row: number, col: number): Scalar;
	/** Materialise a reference (whole columns / rows are clamped to the used range). */
	resolveRef(ref: RefValue): ArrayValue;
	/** Visit the non-blank cells of a reference in row-major order (all sheets of a 3-D reference). */
	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void;
	getFormula(sheet: number, row: number, col: number): string | undefined;
	isRowHidden(sheet: number, row: number): boolean;
	isSubtotalCell(sheet: number, row: number, col: number): boolean;
	usedRange(sheet: number): RangeAddress | null;

	sheetCount(): number;
	sheetName(index: number): string;
	/** Sheet index by name (case-insensitive), -1 when missing. */
	sheetIndex(name: string): number;

	now(): number;
	random(): number;

	/** Defined name visible from `sheet`: the sheet-scoped name first, then the workbook-scoped one. */
	findName(name: string, sheet: number): NameDefinition | undefined;
	/** Table by name (case-insensitive). */
	findTable(name: string): HostTable | undefined;
	/** The table whose range contains a cell. */
	tableAt(sheet: number, row: number, col: number): HostTable | undefined;
	/**
	 * Spill range of the dynamic-array formula anchored at a cell (`A1#`):
	 * the current spill area, or `#REF!` when the cell does not hold a
	 * spilling formula.
	 */
	spillRange(sheet: number, row: number, col: number): RangeAddress | CellError;

	/**
	 * Optional memo for data derived from a reference's values (e.g. a lookup
	 * index), reused while those cells do not change. `key` identifies the
	 * kind of data; `build` computes it.
	 */
	memo?<T>(ref: RefValue, key: string, build: () => T): T;

	/** Optional cell formatting / layout details used by CELL(). */
	readonly cellInfo?: CellInfoProvider;
}

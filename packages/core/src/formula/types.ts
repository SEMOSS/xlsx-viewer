import type { RangeAddress } from '../types/address';
import type { CellError, CellValue } from '../types/cell';
import type { AstNode } from './ast';

/**
 * Runtime values produced while evaluating formulas.
 *
 * - `Scalar` is identical to a cell value: number, string, boolean, error,
 *   or `null` for a blank cell.
 * - `ArrayValue` is an in-memory 2-D array (array constants, array results).
 * - `RefValue` is an unresolved reference to a block of cells; it is only
 *   dereferenced when a function asks for values. Keeping references lazy is
 *   what lets `ROW`, `OFFSET`, `INDEX`, `ISREF`, ... see the reference itself.
 */
export type Scalar = CellValue;

export interface ArrayValue {
	kind: 'array';
	rows: number;
	cols: number;
	/** Row-major values: `values[r][c]`. */
	values: Scalar[][];
}

export interface RefValue {
	kind: 'ref';
	/** Sheet index (into `workbook.sheets`). */
	sheet: number;
	/** Last sheet index of a 3-D reference (`Sheet1:Sheet3!A1`). */
	sheetEnd?: number;
	range: RangeAddress;
}

/** A union of several references (`(A1:A3,C1:C3)`). */
export interface MultiRefValue {
	kind: 'refs';
	refs: RefValue[];
}

/** A LAMBDA closure. */
export interface LambdaValue {
	kind: 'lambda';
	params: string[];
	body: AstNode;
	/** Captured LET / LAMBDA bindings (upper-case names). */
	scope: ReadonlyMap<string, FormulaValue>;
}

export type FormulaValue = Scalar | ArrayValue | RefValue | MultiRefValue | LambdaValue;

/**
 * How the evaluator prepares each argument before calling a function.
 *
 * - `value`: a single value. Single-cell references are dereferenced. When a
 *   multi-cell reference or an array is passed, the function is *lifted*:
 *   it is called once per element (broadcasting across all lifted arguments)
 *   and the results are assembled into an array. In legacy (non-dynamic)
 *   evaluation a multi-cell reference is implicitly intersected instead.
 * - `array`: the whole value. References are materialised into an
 *   `ArrayValue` (whole-column references are clamped to the used range);
 *   scalars become a 1x1 array.
 * - `ref`: the reference itself (`RefValue` / `MultiRefValue`). Other values
 *   are passed through unchanged; the function decides whether that is an error.
 * - `any`: the evaluated value with references left unresolved.
 * - `lazy`: the unevaluated `AstNode`; evaluate it with `ctx.evaluate(node)`.
 */
export type ParamKind = 'value' | 'array' | 'ref' | 'any' | 'lazy';

/** Arguments as delivered to an implementation. Shape depends on the matching ParamKind. */
export type FunctionArg = Scalar | ArrayValue | RefValue | MultiRefValue | LambdaValue | AstNode;

export interface FunctionContext {
	/** Position of the formula being evaluated. */
	readonly sheet: number;
	readonly row: number;
	readonly col: number;
	/** The workbook uses the 1904 date system. */
	readonly date1904: boolean;

	/** Current value of a cell (recalculating it first when necessary). */
	getCellValue(sheet: number, row: number, col: number): Scalar;
	/** Materialise a reference into an array (whole columns / rows are clamped to the used range). */
	resolveRef(ref: RefValue): ArrayValue;
	/**
	 * Visit each *non-blank* cell of a reference in row-major order. Faster
	 * than `resolveRef` for aggregates over large or whole-column ranges.
	 */
	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void;
	/** Formula text (no `=`) of a cell, or `undefined` when it holds a constant. */
	getFormula(sheet: number, row: number, col: number): string | undefined;
	/** Row is hidden (manually or by a filter), for SUBTOTAL / AGGREGATE. */
	isRowHidden(sheet: number, row: number): boolean;
	/** The cell contains a SUBTOTAL or AGGREGATE formula (nested subtotals are skipped). */
	isSubtotalCell(sheet: number, row: number, col: number): boolean;
	/** Bounds of the cells in use on a sheet, or `null` when empty. */
	usedRange(sheet: number): RangeAddress | null;

	sheetCount(): number;
	sheetName(index: number): string;
	/** Sheet index by name (case-insensitive), -1 when missing. */
	sheetIndex(name: string): number;

	/** Evaluate a lazy argument in the current scope. */
	evaluate(node: AstNode): FormulaValue;
	/** Evaluate a lazy argument with extra LET / LAMBDA bindings. */
	evaluateWith(node: AstNode, bindings: ReadonlyMap<string, FormulaValue>): FormulaValue;
	/** Current LET / LAMBDA bindings. */
	readonly scope: ReadonlyMap<string, FormulaValue>;
	/** Invoke a LAMBDA value with arguments. */
	callLambda(fn: LambdaValue, args: FormulaValue[]): FormulaValue;

	/** Current date-time as a serial number (local time). */
	now(): number;
	/** Uniform random number in [0, 1). */
	random(): number;
}

export interface FunctionDefinition {
	/** Canonical upper-case name, including dots (`SUM`, `NORM.S.DIST`). */
	name: string;
	minArgs: number;
	/** Use `Infinity` for variadic functions (Excel's limit of 255 is enforced by the parser). */
	maxArgs: number;
	/** Kind of each parameter; the last entry repeats for any further arguments. */
	params: ParamKind[];
	/** Recalculate on every recalculation (NOW, RAND, OFFSET, INDIRECT, ...). */
	volatile?: boolean;
	/**
	 * When true (the default) and any `value` argument is an error, the
	 * evaluator returns that error without calling `impl`. Set false for
	 * functions that inspect errors (ISERROR, IFERROR, ERROR.TYPE, ...).
	 */
	propagateErrors?: boolean;
	/**
	 * Prefix the function needs inside the file format: `_xlfn.` for Excel
	 * 2010+ functions, `_xlfn._xlws.` for FILTER and SORT.
	 */
	filePrefix?: '_xlfn.' | '_xlfn._xlws.';
	/** The result may be an array that spills (used by the editor to set dynamic-array metadata). */
	returnsArray?: boolean;
	impl: (args: FunctionArg[], ctx: FunctionContext) => FormulaValue;
}

/** Narrowing helpers for runtime values. */
export function isArrayValue(v: unknown): v is ArrayValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'array';
}

export function isRefValue(v: unknown): v is RefValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'ref';
}

export function isMultiRefValue(v: unknown): v is MultiRefValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'refs';
}

export function isLambdaValue(v: unknown): v is LambdaValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'lambda';
}

/** Any reference value (single or union). */
export function isReference(v: unknown): v is RefValue | MultiRefValue {
	return isRefValue(v) || isMultiRefValue(v);
}

/** Build an array value from rows. */
export function makeArray(values: Scalar[][]): ArrayValue {
	return { kind: 'array', rows: values.length, cols: values[0]?.length ?? 0, values };
}

/** 1x1 array wrapping a scalar. */
export function scalarToArray(v: Scalar): ArrayValue {
	return { kind: 'array', rows: 1, cols: 1, values: [[v]] };
}

export type { CellError };

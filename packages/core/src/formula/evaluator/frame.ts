import type { RangeAddress } from '../../types/address';
import type { AstNode } from '../ast';
import type { CellInfoProvider } from '../functions/information/cell';
import type { ArrayMode } from '../invoke';
import type {
	ArrayValue,
	FormulaValue,
	FunctionContext,
	LambdaValue,
	RefValue,
	Scalar,
} from '../types';
import type { EvaluationHost, NameDefinition } from './host';

/** The parts of the evaluator a frame calls back into. */
export interface FrameOwner {
	readonly host: EvaluationHost;
	/** Evaluate a function argument (bare function names become eta-reduced LAMBDAs). */
	evalArg(node: AstNode, frame: EvalFrame): FormulaValue;
	callLambda(fn: LambdaValue, args: readonly FormulaValue[], frame: EvalFrame): FormulaValue;
}

const EMPTY_SCOPE: ReadonlyMap<string, FormulaValue> = new Map();
const NO_NAMES: readonly NameDefinition[] = [];

/**
 * One evaluation scope: the formula cell, the LET / LAMBDA bindings in
 * effect and the array mode. It is the `FunctionContext` handed to
 * function implementations; entering a LET / LAMBDA creates a child frame,
 * everything else reuses the current one.
 */
export class EvalFrame implements FunctionContext, CellInfoProvider {
	constructor(
		readonly owner: FrameOwner,
		readonly sheet: number,
		readonly row: number,
		readonly col: number,
		readonly mode: ArrayMode,
		readonly scope: ReadonlyMap<string, FormulaValue> = EMPTY_SCOPE,
		/**
		 * The frame evaluates a defined name: relative references are relative
		 * to A1 and shift by the formula cell's position.
		 */
		readonly inName = false,
		/** LAMBDA / name nesting depth (recursion guard). */
		readonly depth = 0,
		/** Names being evaluated, innermost last (recursive-name guard). */
		readonly names: readonly NameDefinition[] = NO_NAMES,
	) {}

	get date1904(): boolean {
		return this.owner.host.date1904;
	}

	/** A frame with different bindings, same cell and mode. */
	withScope(scope: ReadonlyMap<string, FormulaValue>, depth = this.depth): EvalFrame {
		return new EvalFrame(
			this.owner,
			this.sheet,
			this.row,
			this.col,
			this.mode,
			scope,
			false,
			depth,
			this.names,
		);
	}

	/** The same scope evaluated with another array mode (array arguments of legacy formulas). */
	withMode(mode: ArrayMode): EvalFrame {
		if (mode === this.mode) {
			return this;
		}
		return new EvalFrame(
			this.owner,
			this.sheet,
			this.row,
			this.col,
			mode,
			this.scope,
			this.inName,
			this.depth,
			this.names,
		);
	}

	/** A frame evaluating a defined name: no LET bindings, references relative to the formula cell. */
	forName(def: NameDefinition): EvalFrame {
		return new EvalFrame(
			this.owner,
			this.sheet,
			this.row,
			this.col,
			this.mode,
			EMPTY_SCOPE,
			true,
			this.depth + 1,
			[...this.names, def],
		);
	}

	getCellValue(sheet: number, row: number, col: number): Scalar {
		return this.owner.host.getCellValue(sheet, row, col);
	}

	resolveRef(ref: RefValue): ArrayValue {
		return this.owner.host.resolveRef(ref);
	}

	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void {
		this.owner.host.forEachCell(ref, visit);
	}

	/**
	 * Memoise data derived from a reference's values (for example an index
	 * for exact-match lookups). The host reuses it within a recalculation
	 * pass until a cell of the reference changes; without host support the
	 * data is simply rebuilt. Duck-typed: `(ctx as LookupMemo).memo?.(...)`.
	 */
	memo<T>(ref: RefValue, key: string, build: () => T): T {
		const host = this.owner.host;
		return host.memo !== undefined ? host.memo(ref, key, build) : build();
	}

	getFormula(sheet: number, row: number, col: number): string | undefined {
		return this.owner.host.getFormula(sheet, row, col);
	}

	isRowHidden(sheet: number, row: number): boolean {
		return this.owner.host.isRowHidden(sheet, row);
	}

	isSubtotalCell(sheet: number, row: number, col: number): boolean {
		return this.owner.host.isSubtotalCell(sheet, row, col);
	}

	usedRange(sheet: number): RangeAddress | null {
		return this.owner.host.usedRange(sheet);
	}

	sheetCount(): number {
		return this.owner.host.sheetCount();
	}

	sheetName(index: number): string {
		return this.owner.host.sheetName(index);
	}

	sheetIndex(name: string): number {
		return this.owner.host.sheetIndex(name);
	}

	evaluate(node: AstNode): FormulaValue {
		return this.owner.evalArg(node, this);
	}

	/**
	 * Evaluate an AST built at run time (INDIRECT's parsed text): its
	 * references are absolute, so no defined-name offset applies.
	 */
	evaluateDetached(node: AstNode): FormulaValue {
		const frame = this.inName ? this.withScope(this.scope) : this;
		return this.owner.evalArg(node, frame);
	}

	evaluateWith(node: AstNode, bindings: ReadonlyMap<string, FormulaValue>): FormulaValue {
		if (bindings.size === 0) {
			return this.owner.evalArg(node, this);
		}
		const scope = new Map(this.scope);
		for (const [k, v] of bindings) {
			scope.set(k.toUpperCase(), v);
		}
		return this.owner.evalArg(node, this.withScope(scope));
	}

	callLambda(fn: LambdaValue, args: FormulaValue[]): FormulaValue {
		return this.owner.callLambda(fn, args, this);
	}

	now(): number {
		return this.owner.host.now();
	}

	random(): number {
		return this.owner.host.random();
	}

	// CELL() support, delegated to the host when it provides the details.

	getNumberFormat(sheet: number, row: number, col: number): string | undefined {
		return this.owner.host.cellInfo?.getNumberFormat?.(sheet, row, col);
	}

	getHorizontalAlignment(sheet: number, row: number, col: number): string | undefined {
		return this.owner.host.cellInfo?.getHorizontalAlignment?.(sheet, row, col);
	}

	isCellLocked(sheet: number, row: number, col: number): boolean {
		return this.owner.host.cellInfo?.isCellLocked?.(sheet, row, col) ?? true;
	}

	getColumnWidth(sheet: number, col: number): { width: number; isDefault: boolean } | undefined {
		return this.owner.host.cellInfo?.getColumnWidth?.(sheet, col);
	}

	getWorkbookName(): string | undefined {
		return this.owner.host.cellInfo?.getWorkbookName?.();
	}
}

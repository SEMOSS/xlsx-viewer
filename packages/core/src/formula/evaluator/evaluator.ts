import type { CellError } from '../../types/cell';
import { cellError, ERRORS, isCellError } from '../../values/errors';
import type {
	ArrayNode,
	AstNode,
	CellRefNode,
	NameNode,
	RangeRefNode,
	SheetQualifier,
	StructuredRefNode,
} from '../ast';
import type { ArrayMode } from '../invoke';
import type { ArrayValue, FormulaValue, LambdaValue, RefValue, Scalar } from '../types';
import { isRefValue } from '../types';
import { callLambdaValue, evalCallExpression, evalFunctionCall, MAX_EVAL_DEPTH } from './calls';
import { EvalFrame, type FrameOwner } from './frame';
import type { EvaluationHost, NameDefinition } from './host';
import { etaLambda, OMITTED } from './lambda';
import { evalBinary, evalPercent, evalUnary } from './operations';
import {
	cellNodeRange,
	makeRef,
	rangeNodeRange,
	resolveSheetSpan,
	type SheetSpan,
} from './references';
import { structuredRange, tableDataRange } from './structured';

/** Value of an array-constant element (the parser only admits literals there). */
function constantValue(node: AstNode): Scalar {
	switch (node.type) {
		case 'number':
		case 'string':
		case 'boolean':
			return node.value;
		case 'error':
			return cellError(node.code);
		default:
			return ERRORS.VALUE;
	}
}

/**
 * Tree-walking formula evaluator. References stay lazy (`RefValue`) until an
 * operator or a function parameter needs values; function arguments are
 * prepared by `invokeFunction` according to each parameter's kind.
 */
export class Evaluator implements FrameOwner {
	private sheetSpans = new WeakMap<SheetQualifier, SheetSpan | CellError>();

	constructor(readonly host: EvaluationHost) {}

	/** Forget cached sheet lookups (after sheets are added, removed, renamed or reordered). */
	reset(): void {
		this.sheetSpans = new WeakMap();
	}

	/**
	 * Evaluate a formula as if it were in cell (row, col) of `sheet`. The
	 * result is raw: references are not dereferenced (see `results.ts`).
	 */
	evaluate(
		ast: AstNode,
		sheet: number,
		row: number,
		col: number,
		mode: ArrayMode = 'dynamic',
	): FormulaValue {
		try {
			return this.eval(ast, new EvalFrame(this, sheet, row, col, mode));
		} catch (e) {
			if (e instanceof RangeError) {
				return ERRORS.NUM;
			}
			throw e;
		}
	}

	eval(node: AstNode, frame: EvalFrame): FormulaValue {
		switch (node.type) {
			case 'number':
			case 'string':
			case 'boolean':
				return node.value;
			case 'error':
				return cellError(node.code);
			case 'missing':
				return null;
			case 'cell':
				return this.cellRef(node, frame);
			case 'range':
				return this.rangeRef(node, frame);
			case 'name':
				return this.name(node, frame);
			case 'structured':
				return this.structured(node, frame);
			case 'function':
				return evalFunctionCall(this, node, frame);
			case 'call':
				return evalCallExpression(this, node, frame);
			case 'unary':
				return evalUnary(this, node, frame);
			case 'percent':
				return evalPercent(this, node.operand, frame);
			case 'binary':
				return evalBinary(this, node, frame);
			case 'array':
				return this.arrayConstant(node);
			case 'spill': {
				const ref = this.cellRef(node.ref, frame);
				return isRefValue(ref) ? this.spillOf(ref) : ref;
			}
		}
	}

	/** Evaluate a function argument: a bare function name becomes an eta-reduced LAMBDA. */
	evalArg(node: AstNode, frame: EvalFrame): FormulaValue {
		if (node.type === 'name' && node.sheet === undefined) {
			const v = this.name(node, frame);
			if (v === ERRORS.NAME) {
				const def = this.host.registry.get(node.name);
				if (def !== undefined) {
					return etaLambda(def);
				}
			}
			return v;
		}
		return this.eval(node, frame);
	}

	callLambda(fn: LambdaValue, args: readonly FormulaValue[], frame: EvalFrame): FormulaValue {
		return callLambdaValue(this, fn, args, frame);
	}

	/** Sheet indices a qualifier denotes (cached per AST node). */
	sheetSpan(q: SheetQualifier | undefined, frame: EvalFrame): SheetSpan | CellError {
		if (q === undefined) {
			return { start: frame.sheet, end: frame.sheet };
		}
		let span = this.sheetSpans.get(q);
		if (span === undefined) {
			span = resolveSheetSpan(q, frame.sheet, this.host);
			this.sheetSpans.set(q, span);
		}
		return span;
	}

	private cellRef(node: CellRefNode, frame: EvalFrame): FormulaValue {
		const span = this.sheetSpan(node.sheet, frame);
		if (isCellError(span)) {
			return span;
		}
		return makeRef(span, cellNodeRange(node, frame.inName, frame.row, frame.col));
	}

	private rangeRef(node: RangeRefNode, frame: EvalFrame): FormulaValue {
		const span = this.sheetSpan(node.sheet, frame);
		if (isCellError(span)) {
			return span;
		}
		return makeRef(span, rangeNodeRange(node, frame.inName, frame.row, frame.col));
	}

	private name(node: NameNode, frame: EvalFrame): FormulaValue {
		if (node.sheet === undefined) {
			const bound = frame.scope.get(node.name.toUpperCase());
			if (bound !== undefined) {
				return bound === OMITTED ? null : bound;
			}
		}
		let sheet = frame.sheet;
		if (node.sheet !== undefined) {
			const span = this.sheetSpan(node.sheet, frame);
			if (isCellError(span)) {
				return span;
			}
			sheet = span.start;
		}
		const def = this.host.findName(node.name, sheet);
		if (def !== undefined) {
			return this.definedName(def, frame);
		}
		if (node.sheet === undefined) {
			const t = this.host.findTable(node.name);
			if (t !== undefined) {
				const range = tableDataRange(t.table);
				return isCellError(range) ? range : { kind: 'ref', sheet: t.sheet, range };
			}
		}
		return ERRORS.NAME;
	}

	/** Value of a defined name, evaluated relative to the formula cell (recursion guarded). */
	definedName(def: NameDefinition, frame: EvalFrame): FormulaValue {
		if (def.ast === null || frame.names.includes(def)) {
			return ERRORS.NAME;
		}
		if (frame.depth >= MAX_EVAL_DEPTH) {
			return ERRORS.NUM;
		}
		return this.eval(def.ast, frame.forName(def));
	}

	private structured(node: StructuredRefNode, frame: EvalFrame): FormulaValue {
		const t =
			node.table !== undefined
				? this.host.findTable(node.table)
				: this.host.tableAt(frame.sheet, frame.row, frame.col);
		if (t === undefined) {
			return ERRORS.REF;
		}
		const range = structuredRange(node, t.table, frame.row);
		return isCellError(range) ? range : { kind: 'ref', sheet: t.sheet, range };
	}

	/** The spill range of the dynamic-array formula at the reference's top-left cell. */
	spillOf(ref: RefValue): FormulaValue {
		if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
			return ERRORS.REF;
		}
		const range = this.host.spillRange(ref.sheet, ref.range.top, ref.range.left);
		return isCellError(range) ? range : { kind: 'ref', sheet: ref.sheet, range };
	}

	private arrayConstant(node: ArrayNode): ArrayValue {
		const values = node.rows.map((row) => row.map(constantValue));
		return { kind: 'array', rows: values.length, cols: values[0]?.length ?? 0, values };
	}
}

import type { CellError } from '../../types/cell';
import { ERRORS, isCellError } from '../../values/errors';
import type { AstNode, CallNode, FunctionNode, SheetQualifier } from '../ast';
import { implicitIntersect, invokeFunction, paramKindAt } from '../invoke';
import type {
	FormulaValue,
	FunctionArg,
	FunctionDefinition,
	LambdaValue,
	RefValue,
} from '../types';
import { isLambdaValue, isMultiRefValue, isRefValue } from '../types';
import type { EvalFrame } from './frame';
import type { EvaluationHost, NameDefinition } from './host';
import { etaBody, etaDefinition, etaScope, OMITTED, paramKey } from './lambda';
import { implicitValue } from './operations';
import type { SheetSpan } from './references';

/** Maximum LAMBDA / name nesting; deeper recursion is `#NUM!` like Excel's recursion limit. */
export const MAX_EVAL_DEPTH = 400;

/** Functions that accept 3-D references (`Sheet1:Sheet3!A1`); others get `#VALUE!`. */
// prettier-ignore
const THREE_D_FUNCTIONS: ReadonlySet<string> = new Set([
	'SUM', 'AVERAGE', 'AVERAGEA', 'COUNT', 'COUNTA', 'MAX', 'MAXA', 'MIN', 'MINA', 'PRODUCT',
	'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA', 'STDEV.S', 'STDEV.P', 'VAR', 'VARA', 'VARP', 'VARPA',
	'VAR.S', 'VAR.P', 'SHEETS', 'ISREF', 'AREAS',
]);

/**
 * Arguments that Excel evaluates as arrays even in legacy (non-CSE)
 * formulas, besides `array` parameters: INDEX's array (`INDEX(A1:A9*2,3)`),
 * LOOKUP's vectors (`LOOKUP(2,1/(A1:A9="x"),B1:B9)`) and the array form of
 * AGGREGATE (`AGGREGATE(14,6,A1:A9/(B1:B9="x"),1)`).
 */
const ARRAY_CONTEXT: ReadonlyMap<string, (index: number) => boolean> = new Map([
	['INDEX', (i: number) => i === 0],
	['LOOKUP', (i: number) => i >= 1],
	['AGGREGATE', (i: number) => i >= 2],
]);

/** What function-call evaluation needs from the evaluator. */
export interface CallEvaluator {
	readonly host: EvaluationHost;
	eval(node: AstNode, frame: EvalFrame): FormulaValue;
	evalArg(node: AstNode, frame: EvalFrame): FormulaValue;
	sheetSpan(q: SheetQualifier | undefined, frame: EvalFrame): SheetSpan | CellError;
	definedName(def: NameDefinition, frame: EvalFrame): FormulaValue;
	spillOf(ref: RefValue): FormulaValue;
}

function is3d(ref: RefValue): boolean {
	return ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet;
}

/** Bind arguments to a LAMBDA's parameters and evaluate its body. */
export function callLambdaValue(
	ev: CallEvaluator,
	fn: LambdaValue,
	args: readonly FormulaValue[],
	frame: EvalFrame,
): FormulaValue {
	if (frame.depth >= MAX_EVAL_DEPTH) {
		return ERRORS.NUM;
	}
	const def = etaDefinition(fn);
	if (def !== undefined) {
		if (args.length < def.minArgs || args.length > def.maxArgs) {
			return ERRORS.VALUE;
		}
		return ev.eval(etaBody(def, args.length), frame.withScope(etaScope(args), frame.depth + 1));
	}
	if (args.length > fn.params.length) {
		return ERRORS.VALUE;
	}
	const scope = new Map(fn.scope);
	for (let i = 0; i < fn.params.length; i++) {
		const p = fn.params[i];
		const optional = p.charCodeAt(0) === 91; /* [ */
		if (i >= args.length && !optional) {
			return ERRORS.VALUE;
		}
		scope.set(paramKey(p), i < args.length ? args[i] : OMITTED);
	}
	return ev.eval(fn.body, frame.withScope(scope, frame.depth + 1));
}

function invokeLambda(
	ev: CallEvaluator,
	fn: LambdaValue,
	argNodes: readonly AstNode[],
	frame: EvalFrame,
): FormulaValue {
	const args = argNodes.map((a) => (a.type === 'missing' ? OMITTED : ev.evalArg(a, frame)));
	return callLambdaValue(ev, fn, args, frame);
}

/** Prepare the arguments of a built-in function and invoke it through `invokeFunction`. */
function invokeDefinition(
	ev: CallEvaluator,
	def: FunctionDefinition,
	args: readonly AstNode[],
	frame: EvalFrame,
): FormulaValue {
	const n = args.length;
	if (n < def.minArgs || n > def.maxArgs) {
		return ERRORS.VALUE;
	}
	const raw: FunctionArg[] = new Array(n);
	const legacy = frame.mode === 'legacy';
	const arrayContext = legacy ? ARRAY_CONTEXT.get(def.name) : undefined;
	for (let i = 0; i < n; i++) {
		const arg = args[i];
		const kind = paramKindAt(def, i);
		if (kind === 'lazy' || arg.type === 'missing') {
			raw[i] = arg;
			continue;
		}
		if (kind === 'value' && arg.type === 'cell' && !frame.inName) {
			// Fast path: a plain cell argument is read directly.
			const span = ev.sheetSpan(arg.sheet, frame);
			if (!isCellError(span) && span.start === span.end) {
				raw[i] = ev.host.getCellValue(span.start, arg.ref.row, arg.ref.col);
				continue;
			}
		}
		// Array parameters evaluate their argument as an array formula, even in legacy formulas
		// (`SUMPRODUCT(1/COUNTIF(A1:A9,A1:A9))` needs no Ctrl+Shift+Enter in Excel).
		const asArray = legacy && (kind === 'array' || arrayContext?.(i) === true);
		let v = ev.evalArg(arg, asArray ? frame.withMode('dynamic') : frame);
		if (isRefValue(v)) {
			if (is3d(v) && !THREE_D_FUNCTIONS.has(def.name)) {
				return ERRORS.VALUE;
			}
			if (legacy && kind === 'value' && !isSingleCell(v)) {
				v = implicitIntersect(v, frame);
			}
		}
		raw[i] = v;
	}
	// Legacy formulas intersect references (above) but still evaluate array values element-wise,
	// as Excel does for `SUM(COUNTIF(A:A,{"a","b"}))`; their cell shows the first element.
	return invokeFunction(def, raw, frame, 'dynamic');
}

function isSingleCell(ref: RefValue): boolean {
	return ref.range.top === ref.range.bottom && ref.range.left === ref.range.right;
}

/**
 * A function call: a LET / LAMBDA-bound name, the evaluator's own
 * ANCHORARRAY / SINGLE, a built-in function, a defined name holding a
 * LAMBDA, or `#NAME?`.
 */
export function evalFunctionCall(
	ev: CallEvaluator,
	node: FunctionNode,
	frame: EvalFrame,
): FormulaValue {
	const name = node.name;
	const local = frame.scope.get(name);
	if (local !== undefined) {
		return isLambdaValue(local) ? invokeLambda(ev, local, node.args, frame) : ERRORS.VALUE;
	}
	if (name === 'ANCHORARRAY' || name === 'SINGLE') {
		if (node.args.length !== 1) {
			return ERRORS.VALUE;
		}
		const v = ev.eval(node.args[0], frame);
		if (name === 'SINGLE') {
			return implicitValue(v, frame);
		}
		const ref = isRefValue(v) ? v : isMultiRefValue(v) && v.refs.length === 1 ? v.refs[0] : null;
		return ref !== null ? ev.spillOf(ref) : isCellError(v) ? v : ERRORS.REF;
	}
	const def = ev.host.registry.get(name);
	if (def !== undefined) {
		return invokeDefinition(ev, def, node.args, frame);
	}
	const named = ev.host.findName(name, frame.sheet);
	if (named !== undefined) {
		const fn = ev.definedName(named, frame);
		if (isLambdaValue(fn)) {
			return invokeLambda(ev, fn, node.args, frame);
		}
		return isCellError(fn) ? fn : ERRORS.VALUE;
	}
	return ERRORS.NAME;
}

/** Calling the value of an expression: `LAMBDA(x, x + 1)(2)`. */
export function evalCallExpression(
	ev: CallEvaluator,
	node: CallNode,
	frame: EvalFrame,
): FormulaValue {
	const callee = ev.evalArg(node.callee, frame);
	if (isLambdaValue(callee)) {
		return invokeLambda(ev, callee, node.args, frame);
	}
	return isCellError(callee) ? callee : ERRORS.VALUE;
}

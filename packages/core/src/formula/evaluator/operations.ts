import type { CellError } from '../../types/cell';
import { ERRORS, isCellError } from '../../values/errors';
import type { AstNode, BinaryNode, SheetQualifier, UnaryNode } from '../ast';
import { implicitIntersect } from '../invoke';
import type { ArrayValue, FormulaValue, RefValue, Scalar } from '../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../types';
import type { EvalFrame } from './frame';
import type { EvaluationHost } from './host';
import { OMITTED } from './lambda';
import { applyOperator, broadcastBinary, mapArray, negate, percent } from './operators';
import {
	intersectionOperator,
	intersectRef,
	rangeOperator,
	type SheetSpan,
	unionOperator,
} from './references';

/** What operator evaluation needs from the evaluator. */
export interface OperandEvaluator {
	readonly host: EvaluationHost;
	eval(node: AstNode, frame: EvalFrame): FormulaValue;
	sheetSpan(q: SheetQualifier | undefined, frame: EvalFrame): SheetSpan | CellError;
}

function isOperandArray(v: Scalar | ArrayValue): v is ArrayValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'array';
}

function is3d(ref: RefValue): boolean {
	return ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet;
}

/** Reduce a value for an operator: references are read (implicitly intersected in legacy mode). */
export function toOperand(
	ev: OperandEvaluator,
	v: FormulaValue,
	frame: EvalFrame,
): Scalar | ArrayValue {
	if (isRefValue(v)) {
		if (is3d(v)) {
			return ERRORS.VALUE;
		}
		const r = v.range;
		if (r.top === r.bottom && r.left === r.right) {
			return ev.host.getCellValue(v.sheet, r.top, r.left);
		}
		if (frame.mode === 'legacy') {
			return implicitIntersect(v, frame);
		}
		return toOperand(ev, ev.host.resolveRef(v), frame);
	}
	if (isArrayValue(v)) {
		if (v.rows === 1 && v.cols === 1) {
			return v.values[0][0];
		}
		return v.rows === 0 || v.cols === 0 ? ERRORS.CALC : v;
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? toOperand(ev, v.refs[0], frame) : ERRORS.VALUE;
	}
	if (isLambdaValue(v)) {
		return ERRORS.CALC;
	}
	if (v === OMITTED) {
		return null;
	}
	return v;
}

/** Evaluate an operator operand down to a scalar or an array (plain cells are read directly). */
export function evalOperand(
	ev: OperandEvaluator,
	node: AstNode,
	frame: EvalFrame,
): Scalar | ArrayValue {
	switch (node.type) {
		case 'number':
		case 'string':
		case 'boolean':
			return node.value;
		case 'cell':
			if (!frame.inName) {
				const span = ev.sheetSpan(node.sheet, frame);
				if (isCellError(span)) {
					return span;
				}
				if (span.start === span.end) {
					return ev.host.getCellValue(span.start, node.ref.row, node.ref.col);
				}
			}
			break;
	}
	return toOperand(ev, ev.eval(node, frame), frame);
}

/** `@`: implicit intersection, keeping references as references. */
export function implicitValue(v: FormulaValue, frame: EvalFrame): FormulaValue {
	if (isRefValue(v)) {
		return intersectRef(v, frame.row, frame.col);
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? intersectRef(v.refs[0], frame.row, frame.col) : ERRORS.VALUE;
	}
	if (isArrayValue(v)) {
		return v.rows > 0 && v.cols > 0 ? v.values[0][0] : ERRORS.CALC;
	}
	return v;
}

function unaryValue(
	ev: OperandEvaluator,
	operand: AstNode,
	frame: EvalFrame,
	f: (v: Scalar, d: boolean) => Scalar,
): FormulaValue {
	const v = evalOperand(ev, operand, frame);
	const d = ev.host.date1904;
	return isOperandArray(v) ? mapArray(v, (x) => f(x, d)) : f(v, d);
}

/** Unary `+` (returns its operand unchanged, references included), `-` and `@`. */
export function evalUnary(ev: OperandEvaluator, node: UnaryNode, frame: EvalFrame): FormulaValue {
	if (node.op === '+') {
		return ev.eval(node.operand, frame);
	}
	if (node.op === '@') {
		return implicitValue(ev.eval(node.operand, frame), frame);
	}
	return unaryValue(ev, node.operand, frame, negate);
}

/** Postfix `%`. */
export function evalPercent(
	ev: OperandEvaluator,
	operand: AstNode,
	frame: EvalFrame,
): FormulaValue {
	return unaryValue(ev, operand, frame, percent);
}

/** Binary operators: reference operators on references, the others element-wise on values. */
export function evalBinary(ev: OperandEvaluator, node: BinaryNode, frame: EvalFrame): FormulaValue {
	const op = node.op;
	switch (op) {
		case ':':
			return rangeOperator(ev.eval(node.left, frame), ev.eval(node.right, frame));
		case ',':
			return unionOperator(ev.eval(node.left, frame), ev.eval(node.right, frame));
		case ' ':
			return intersectionOperator(ev.eval(node.left, frame), ev.eval(node.right, frame));
	}
	const a = evalOperand(ev, node.left, frame);
	const b = evalOperand(ev, node.right, frame);
	const d = ev.host.date1904;
	if (!isOperandArray(a) && !isOperandArray(b)) {
		return applyOperator(op, a, b, d);
	}
	return broadcastBinary(a, b, (x, y) => applyOperator(op, x, y, d));
}

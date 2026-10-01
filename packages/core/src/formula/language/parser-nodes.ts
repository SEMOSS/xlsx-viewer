import type { AstNode, CellRefNode, ErrorNode, NameNode, RangeRefNode } from '../ast';
import { stripFilePrefixes } from './names';
import { FormulaParseError } from './parse-error';
import { matchErrorCode } from './scan';
import type { FormulaToken } from './tokens';

/** AST builders and grammar tables shared by the parser. */

/** Precedence of arithmetic / text / comparison operators (reference operators bind tighter). */
export function binaryPrecedence(op: string): number {
	switch (op) {
		case '^':
			return 5;
		case '*':
		case '/':
			return 4;
		case '+':
		case '-':
			return 3;
		case '&':
			return 2;
		case '=':
		case '<>':
		case '<':
		case '>':
		case '<=':
		case '>=':
			return 1;
		default:
			return -1;
	}
}

/** Precedence of the reference operators, which bind tighter than every other operator. */
export const REFERENCE_OPERATORS = { ':': 10, ' ': 9, ',': 8 } as const;

/** The reference operator a token stands for (`:` range, whitespace intersection, `,` union). */
export function referenceOperator(t: FormulaToken): keyof typeof REFERENCE_OPERATORS | null {
	if (t.type === 'intersect') {
		return ' ';
	}
	if (t.type === 'union') {
		return ',';
	}
	return t.type === 'operator' && t.text === ':' ? ':' : null;
}

/** Error for a token the grammar does not allow here (`undefined`: the formula ended early). */
export function unexpectedToken(t: FormulaToken | undefined, end: number): FormulaParseError {
	if (t === undefined) {
		return new FormulaParseError('Unexpected end of formula', end);
	}
	if (t.incomplete) {
		const what =
			t.type === 'string'
				? 'text value'
				: t.type === 'structured'
					? 'table reference'
					: 'reference';
		return new FormulaParseError(`Unterminated ${what} '${t.text}'`, t.start);
	}
	return new FormulaParseError(`Unexpected '${t.text}'`, t.start);
}

/** Value of a string literal token (quotes removed, `""` un-doubled). */
export function decodeString(text: string): string {
	const body = text.slice(1, -1);
	return body.indexOf('""') < 0 ? body : body.replace(/""/g, '"');
}

/** Cell or range node of a `ref` token. */
export function refNode(t: FormulaToken): CellRefNode | RangeRefNode {
	const area = t.area!;
	const node: CellRefNode | RangeRefNode =
		area.kind === 'cell'
			? { type: 'cell', ref: area.start }
			: { type: 'range', kind: area.kind, start: area.start, end: area.end };
	if (t.sheet !== undefined) {
		node.sheet = t.sheet;
	}
	return node;
}

/** Error node of an `error` token (`#N/A`, `Sheet1!#REF!`, dead `#REF!A1`). */
export function errorNode(t: FormulaToken): ErrorNode {
	const node: ErrorNode = { type: 'error', code: matchErrorCode(t.text, t.qualifierLength ?? 0)! };
	if (t.sheet !== undefined) {
		node.sheet = t.sheet;
	}
	return node;
}

/** Name node of a `name` token; `_xlpm.` and friends are dropped from unqualified names. */
export function nameNode(t: FormulaToken): NameNode {
	if (t.sheet === undefined) {
		return { type: 'name', name: stripFilePrefixes(t.text) };
	}
	return { type: 'name', name: t.text.slice(t.qualifierLength), sheet: t.sheet };
}

/** Nodes that may produce a reference, so they can be operands of `:`, ` ` and `,`. */
export function isReferenceOperand(node: AstNode): boolean {
	switch (node.type) {
		case 'cell':
		case 'range':
		case 'name':
		case 'structured':
		case 'function':
		case 'call':
		case 'spill':
			return true;
		case 'error':
			return node.code === '#REF!';
		case 'binary':
			return node.op === ':' || node.op === ' ' || node.op === ',';
		case 'unary':
			return node.op === '@';
		default:
			return false;
	}
}

/** Nodes that can be called with an argument list (`LAMBDA(x,x)(1)`). */
export function isCallable(node: AstNode): boolean {
	return node.type === 'function' || node.type === 'call' || node.type === 'name';
}

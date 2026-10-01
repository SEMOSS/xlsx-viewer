import type { AstNode, BinaryOperator, NumberNode } from '../ast';
import { normalizeFunctionName } from './names';
import { FormulaParseError } from './parse-error';
import {
	binaryPrecedence,
	decodeString,
	errorNode,
	isCallable,
	isReferenceOperand,
	nameNode,
	REFERENCE_OPERATORS,
	referenceOperator,
	refNode,
	unexpectedToken,
} from './parser-nodes';
import { matchErrorCode } from './scan';
import { parseStructuredReference } from './structured';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken, TokenizeOptions } from './tokens';

/** Excel accepts at most 255 arguments per function call. */
export const MAX_FUNCTION_ARGS = 255;

/** Nesting limit for parentheses, calls and arrays (Excel allows 64 nested functions). */
const MAX_NESTING = 256;

export type ParseOptions = TokenizeOptions;

/**
 * Parse a formula (leading `=` optional) into an AST.
 *
 * Precedence follows Excel, from tightest to loosest: `:` (range), ` `
 * (intersection), `,` (union), unary `-` `+` `@`, `%`, `^`, `* /`, `+ -`,
 * `&`, comparisons. Unary minus binds tighter than `^` (`=-2^2` is 4) and
 * `^` is left-associative. Function names are canonical upper case with the
 * file prefixes removed; `_xlfn.ANCHORARRAY(A1)` becomes a spill reference,
 * `_xlfn.SINGLE(x)` the `@` operator and `_xlpm.` is dropped from names.
 *
 * @throws FormulaParseError when the text is not a valid formula.
 */
export function parseFormula(formula: string, options?: ParseOptions): AstNode {
	return new Parser(tokenizeFormula(formula, options), formula.length).parse();
}

/** Parse an already tokenized formula (see {@link tokenizeFormula}). */
export function parseFormulaTokens(tokens: readonly FormulaToken[]): AstNode {
	const end = tokens.length > 0 ? tokens[tokens.length - 1].end : 0;
	return new Parser(tokens, end).parse();
}

class Parser {
	/** Significant tokens (whitespace and the leading `=` removed). */
	private readonly toks: FormulaToken[] = [];
	private i = 0;
	private depth = 0;

	constructor(
		tokens: readonly FormulaToken[],
		private readonly end: number,
	) {
		for (const t of tokens) {
			if (t.type !== 'whitespace' && t.type !== 'equals') {
				this.toks.push(t);
			}
		}
	}

	parse(): AstNode {
		if (this.toks.length === 0) {
			throw new FormulaParseError('The formula is empty', this.end);
		}
		const node = this.expression(0);
		if (this.i < this.toks.length) {
			throw this.unexpected(this.toks[this.i]);
		}
		return node;
	}

	private unexpected(t: FormulaToken | undefined): FormulaParseError {
		return unexpectedToken(t, this.end);
	}

	private enter(t: FormulaToken): void {
		if (++this.depth > MAX_NESTING) {
			throw new FormulaParseError('The formula is nested too deeply', t.start);
		}
	}

	/** Binary operators of at least `minPrec` (left-associative precedence climbing). */
	private expression(minPrec: number): AstNode {
		let left = this.prefixed();
		for (;;) {
			const t = this.toks[this.i];
			if (t === undefined || t.type !== 'operator') {
				return left;
			}
			const prec = binaryPrecedence(t.text);
			if (prec < 0) {
				throw this.unexpected(t);
			}
			if (prec < minPrec) {
				return left;
			}
			this.i++;
			const right = this.expression(prec + 1);
			left = { type: 'binary', op: t.text as BinaryOperator, left, right };
		}
	}

	/** Unary operators, then postfix `%` (negation binds tighter than `%` and `^`). */
	private prefixed(): AstNode {
		const first = this.i;
		while (this.i < this.toks.length && this.toks[this.i].type === 'prefix') {
			this.i++;
		}
		const opsEnd = this.i;
		let node = this.referenceExpression(REFERENCE_OPERATORS[',']);
		for (let k = opsEnd - 1; k >= first; k--) {
			node = { type: 'unary', op: this.toks[k].text as '+' | '-' | '@', operand: node };
		}
		while (this.i < this.toks.length && this.toks[this.i].type === 'percent') {
			this.i++;
			node = { type: 'percent', operand: node };
		}
		return node;
	}

	/** Reference operators: `:` binds tighter than intersection, which binds tighter than union. */
	private referenceExpression(minPrec: number): AstNode {
		let left = this.postfix();
		for (;;) {
			const t = this.toks[this.i];
			const op = t === undefined ? null : referenceOperator(t);
			if (op === null || REFERENCE_OPERATORS[op] < minPrec) {
				return left;
			}
			this.i++;
			const right = this.referenceExpression(REFERENCE_OPERATORS[op] + 1);
			if (!isReferenceOperand(left) || !isReferenceOperand(right)) {
				const name = op === ':' ? 'range' : op === ',' ? 'union' : 'intersection';
				throw new FormulaParseError(
					`The ${name} operator needs a reference on each side`,
					t!.start,
				);
			}
			left = { type: 'binary', op, left, right };
		}
	}

	/** A primary followed by the spill operator `#`. */
	private postfix(): AstNode {
		const node = this.primary();
		const t = this.toks[this.i];
		if (t === undefined || t.type !== 'spill') {
			return node;
		}
		this.i++;
		if (node.type === 'cell') {
			return { type: 'spill', ref: node };
		}
		if (node.type === 'name') {
			return { type: 'function', name: 'ANCHORARRAY', args: [node] };
		}
		throw new FormulaParseError(
			'The spill operator # must follow a single cell reference',
			t.start,
		);
	}

	private primary(): AstNode {
		const t = this.toks[this.i];
		if (t === undefined || t.incomplete) {
			throw this.unexpected(t);
		}
		switch (t.type) {
			case 'number': {
				const value = Number(t.text);
				if (!Number.isFinite(value)) {
					throw new FormulaParseError(`Number out of range: ${t.text}`, t.start);
				}
				this.i++;
				return { type: 'number', value };
			}
			case 'string':
				this.i++;
				return { type: 'string', value: decodeString(t.text) };
			case 'boolean':
				this.i++;
				return { type: 'boolean', value: t.text.toUpperCase() === 'TRUE' };
			case 'error':
				this.i++;
				return errorNode(t);
			case 'ref':
				this.i++;
				return refNode(t);
			case 'name':
				this.i++;
				return nameNode(t);
			case 'structured': {
				const node = parseStructuredReference(t.text);
				if (node === null) {
					throw new FormulaParseError(`Invalid table reference '${t.text}'`, t.start);
				}
				this.i++;
				return node;
			}
			case 'function':
				return this.calls(this.functionCall(t));
			case 'paren-open':
				return this.calls(this.group(t));
			case 'array-open':
				return this.array(t);
			default:
				throw this.unexpected(t);
		}
	}

	private functionCall(t: FormulaToken): AstNode {
		this.i++;
		const name = normalizeFunctionName(t.text);
		const args = this.args();
		if (args.length === 1) {
			const arg = args[0];
			if (name === 'ANCHORARRAY' && arg.type === 'cell' && arg.parens === undefined) {
				return { type: 'spill', ref: arg };
			}
			if (name === 'SINGLE' && arg.type !== 'missing') {
				return { type: 'unary', op: '@', operand: arg };
			}
		}
		return { type: 'function', name, args };
	}

	/** Argument list starting at a `call-open` token; omitted arguments become `missing` nodes. */
	private args(): AstNode[] {
		const open = this.toks[this.i];
		if (open === undefined || open.type !== 'call-open') {
			throw this.unexpected(open);
		}
		this.enter(open);
		this.i++;
		const args: AstNode[] = [];
		if (this.toks[this.i]?.type === 'call-close') {
			this.i++;
			this.depth--;
			return args;
		}
		for (;;) {
			const t = this.toks[this.i];
			if (t === undefined) {
				throw new FormulaParseError('Missing closing parenthesis', this.end);
			}
			args.push(
				t.type === 'arg-sep' || t.type === 'call-close' ? { type: 'missing' } : this.expression(0),
			);
			const sep = this.toks[this.i];
			if (sep === undefined) {
				throw new FormulaParseError('Missing closing parenthesis', this.end);
			}
			if (sep.type === 'call-close') {
				this.i++;
				break;
			}
			if (sep.type !== 'arg-sep') {
				throw this.unexpected(sep);
			}
			if (args.length === MAX_FUNCTION_ARGS) {
				throw new FormulaParseError(
					`A function accepts at most ${MAX_FUNCTION_ARGS} arguments`,
					sep.start,
				);
			}
			this.i++;
		}
		this.depth--;
		return args;
	}

	/** Calls applied to an expression: `LAMBDA(x, x + 1)(2)`. */
	private calls(callee: AstNode): AstNode {
		let node = callee;
		while (this.toks[this.i]?.type === 'call-open') {
			if (!isCallable(node)) {
				throw new FormulaParseError(
					'Only functions and LAMBDA values can be called',
					this.toks[this.i].start,
				);
			}
			node = { type: 'call', callee: node, args: this.args() };
		}
		return node;
	}

	private group(open: FormulaToken): AstNode {
		this.enter(open);
		this.i++;
		const inner = this.expression(0);
		const close = this.toks[this.i];
		if (close === undefined) {
			throw new FormulaParseError('Missing closing parenthesis', this.end);
		}
		if (close.type !== 'paren-close') {
			throw this.unexpected(close);
		}
		this.i++;
		this.depth--;
		inner.parens = (inner.parens ?? 0) + 1;
		return inner;
	}

	private array(open: FormulaToken): AstNode {
		this.enter(open);
		this.i++;
		const rows: AstNode[][] = [];
		let row: AstNode[] = [];
		for (;;) {
			row.push(this.arrayElement());
			const t = this.toks[this.i];
			if (t === undefined) {
				throw new FormulaParseError('Missing closing brace in array constant', this.end);
			}
			if (t.type !== 'array-col-sep' && t.type !== 'array-row-sep' && t.type !== 'array-close') {
				throw this.unexpected(t);
			}
			this.i++;
			if (t.type === 'array-col-sep') {
				continue;
			}
			rows.push(row);
			if (t.type === 'array-close') {
				break;
			}
			row = [];
		}
		const width = rows[0].length;
		for (const r of rows) {
			if (r.length !== width) {
				throw new FormulaParseError(
					'Every row of an array constant needs the same number of values',
					open.start,
				);
			}
		}
		this.depth--;
		return { type: 'array', rows };
	}

	/** Array constants hold numbers (optionally signed), text, logical values and error literals. */
	private arrayElement(): AstNode {
		let t = this.toks[this.i];
		let negate = false;
		if (t !== undefined && t.type === 'prefix' && t.text !== '@') {
			negate = t.text === '-';
			this.i++;
			t = this.toks[this.i];
			if (t !== undefined && t.type !== 'number') {
				throw new FormulaParseError('Only numbers can be signed in an array constant', t.start);
			}
		}
		if (t === undefined) {
			throw this.unexpected(t);
		}
		const plainError =
			t.type === 'error' &&
			t.sheet === undefined &&
			matchErrorCode(t.text, 0)!.length === t.text.length;
		if (t.type !== 'number' && t.type !== 'string' && t.type !== 'boolean' && !plainError) {
			throw new FormulaParseError(
				'Array constants can only contain numbers, text, logical values and error values',
				t.start,
			);
		}
		const node = this.primary();
		if (negate) {
			const num = node as NumberNode;
			num.value = num.value === 0 ? 0 : -num.value;
		}
		return node;
	}
}

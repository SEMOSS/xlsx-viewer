import type { SheetQualifier } from '../ast';
import {
	CH_AMP,
	CH_AT,
	CH_CARET,
	CH_COLON,
	CH_COMMA,
	CH_CR,
	CH_DOLLAR,
	CH_DOT,
	CH_DQUOTE,
	CH_EQ,
	CH_GT,
	CH_HASH,
	CH_LBRACE,
	CH_LBRACKET,
	CH_LF,
	CH_LPAREN,
	CH_LT,
	CH_MINUS,
	CH_PERCENT,
	CH_PLUS,
	CH_RBRACE,
	CH_RPAREN,
	CH_SEMI,
	CH_SLASH,
	CH_SPACE,
	CH_SQUOTE,
	CH_STAR,
	CH_TAB,
	isDigit,
	isNameStart,
	isSpace,
} from './chars';
import {
	deadReferenceEnd,
	type LexerSink,
	lexBracket,
	lexQuotedQualifier,
	lexWord,
} from './lexer-words';
import { matchErrorCode, scanA1Area, type ScannedArea } from './scan';
import type { FormulaToken, FormulaTokenType, TokenizeOptions } from './tokens';
import { endsOperand, endsReferenceOperand, startsReferenceOperand } from './tokens';

/** Error tokens only join intersections as dead references (`#REF!`, `Sheet1!#REF!`). */
function isReferenceError(t: FormulaToken): boolean {
	if (t.type !== 'error') {
		return true;
	}
	const q = t.qualifierLength ?? 0;
	return t.text.slice(q, q + 5).toUpperCase() === '#REF!';
}

const FRAME_CALL = 1;
const FRAME_GROUP = 2;
const FRAME_ARRAY = 3;

/**
 * Split a formula (with or without the leading `=`) into tokens.
 *
 * The tokenizer never throws: characters the grammar does not allow become
 * `unknown` tokens and unterminated strings / sheet names / structured
 * references are flagged `incomplete`, so it is safe on partial input typed
 * in the formula bar. Joining the token texts reproduces the input exactly.
 *
 * Context decides what `,` means (argument separator inside a function call,
 * column separator inside an array constant, union elsewhere), whether `+` /
 * `-` are unary, and whether whitespace between two references is the
 * intersection operator.
 */
export function tokenizeFormula(formula: string, options?: TokenizeOptions): FormulaToken[] {
	return new Lexer(formula, options?.r1c1).run();
}

class Lexer implements LexerSink {
	private readonly tokens: FormulaToken[] = [];
	/** Open brackets: function call, group or array constant. */
	private readonly frames: number[] = [];
	/** Last significant (non-whitespace) token. */
	private last: FormulaToken | undefined;

	constructor(
		readonly src: string,
		readonly base: { row: number; col: number } | undefined,
	) {}

	run(): FormulaToken[] {
		const src = this.src;
		let p = 0;
		if (src.charCodeAt(0) === CH_EQ) {
			this.tokens.push({ type: 'equals', text: '=', start: 0, end: 1 });
			p = 1;
		}
		while (p < src.length) {
			p = this.step(p);
		}
		return this.tokens;
	}

	private top(): number | undefined {
		return this.frames[this.frames.length - 1];
	}

	/** Append a token; whitespace between two reference operands becomes an intersection. */
	private add(token: FormulaToken): void {
		const tokens = this.tokens;
		if (token.type !== 'whitespace') {
			const prev = tokens[tokens.length - 1];
			if (
				prev !== undefined &&
				prev.type === 'whitespace' &&
				this.last !== undefined &&
				endsReferenceOperand(this.last.type) &&
				startsReferenceOperand(token.type) &&
				isReferenceError(this.last) &&
				isReferenceError(token) &&
				this.top() !== FRAME_ARRAY
			) {
				prev.type = 'intersect';
			}
			this.last = token;
		}
		tokens.push(token);
	}

	emit(type: FormulaTokenType, start: number, end: number): number {
		this.add({ type, text: this.src.slice(start, end), start, end });
		return end;
	}

	emitQualified(
		type: FormulaTokenType,
		start: number,
		end: number,
		scanned: ScannedArea | null,
		sheet?: SheetQualifier,
		qualifierLength?: number,
	): number {
		const token: FormulaToken = { type, text: this.src.slice(start, end), start, end };
		if (sheet !== undefined) {
			token.sheet = sheet;
			token.qualifierLength = qualifierLength;
		}
		if (scanned !== null) {
			token.area = scanned.area;
		}
		this.add(token);
		return end;
	}

	emitIncomplete(type: FormulaTokenType, start: number): number {
		const end = this.src.length;
		this.add({ type, text: this.src.slice(start), start, end, incomplete: true });
		return end;
	}

	unknown(start: number, end: number, incomplete = false): number {
		const token: FormulaToken = { type: 'unknown', text: this.src.slice(start, end), start, end };
		if (incomplete) {
			token.incomplete = true;
		}
		this.add(token);
		return end;
	}

	private step(p: number): number {
		const src = this.src;
		const c = src.charCodeAt(p);
		switch (c) {
			case CH_SPACE:
			case CH_TAB:
			case CH_LF:
			case CH_CR: {
				let e = p + 1;
				while (e < src.length && isSpace(src.charCodeAt(e))) {
					e++;
				}
				return this.emit('whitespace', p, e);
			}
			case CH_DQUOTE:
				return this.string(p);
			case CH_SQUOTE:
				return lexQuotedQualifier(this, p);
			case CH_LBRACKET:
				return lexBracket(this, p);
			case CH_LBRACE:
				this.frames.push(FRAME_ARRAY);
				return this.emit('array-open', p, p + 1);
			case CH_RBRACE:
				if (this.top() === FRAME_ARRAY) {
					this.frames.pop();
				}
				return this.emit('array-close', p, p + 1);
			case CH_LPAREN:
				return this.openParen(p);
			case CH_RPAREN: {
				const top = this.top();
				if (top === FRAME_CALL || top === FRAME_GROUP) {
					this.frames.pop();
				}
				return this.emit(top === FRAME_CALL ? 'call-close' : 'paren-close', p, p + 1);
			}
			case CH_COMMA: {
				const top = this.top();
				const type =
					top === FRAME_CALL ? 'arg-sep' : top === FRAME_ARRAY ? 'array-col-sep' : 'union';
				return this.emit(type, p, p + 1);
			}
			case CH_SEMI:
				return this.emit(this.top() === FRAME_ARRAY ? 'array-row-sep' : 'unknown', p, p + 1);
			case CH_PLUS:
			case CH_MINUS: {
				const binary = this.last !== undefined && endsOperand(this.last.type);
				return this.emit(binary ? 'operator' : 'prefix', p, p + 1);
			}
			case CH_STAR:
			case CH_SLASH:
			case CH_CARET:
			case CH_AMP:
			case CH_EQ:
			case CH_COLON:
				return this.emit('operator', p, p + 1);
			case CH_LT: {
				const n = src.charCodeAt(p + 1);
				return this.emit('operator', p, n === CH_GT || n === CH_EQ ? p + 2 : p + 1);
			}
			case CH_GT:
				return this.emit('operator', p, src.charCodeAt(p + 1) === CH_EQ ? p + 2 : p + 1);
			case CH_PERCENT:
				return this.emit('percent', p, p + 1);
			case CH_AT:
				return this.emit('prefix', p, p + 1);
			case CH_HASH:
				return this.hash(p);
			case CH_DOT:
				return isDigit(src.charCodeAt(p + 1)) ? this.number(p) : this.unknown(p, p + 1);
			case CH_DOLLAR:
				return lexWord(this, p);
			default:
				if (isDigit(c)) {
					return this.number(p);
				}
				if (isNameStart(c)) {
					return lexWord(this, p);
				}
				return this.unknown(p, p + 1);
		}
	}

	private string(s: number): number {
		const src = this.src;
		let i = s + 1;
		while (i < src.length) {
			if (src.charCodeAt(i) === CH_DQUOTE) {
				if (src.charCodeAt(i + 1) !== CH_DQUOTE) {
					return this.emit('string', s, i + 1);
				}
				i++;
			}
			i++;
		}
		return this.emitIncomplete('string', s);
	}

	private number(s: number): number {
		const src = this.src;
		let i = s;
		while (isDigit(src.charCodeAt(i))) {
			i++;
		}
		if (i > s && src.charCodeAt(i) === CH_COLON && this.base === undefined) {
			// Whole rows: `1:3`.
			const rows = scanA1Area(src, s);
			if (rows !== null) {
				return this.emitQualified('ref', s, rows.end, rows);
			}
		}
		if (src.charCodeAt(i) === CH_DOT) {
			i++;
			while (isDigit(src.charCodeAt(i))) {
				i++;
			}
		}
		if ((src.charCodeAt(i) | 32) === 101 /* e */) {
			let j = i + 1;
			const sign = src.charCodeAt(j);
			if (sign === CH_PLUS || sign === CH_MINUS) {
				j++;
			}
			if (isDigit(src.charCodeAt(j))) {
				while (isDigit(src.charCodeAt(j))) {
					j++;
				}
				i = j;
			}
		}
		return this.emit('number', s, i);
	}

	/** `(` right after a function name, `)` of a call or a group calls; anything else groups. */
	private openParen(p: number): number {
		const last = this.last;
		const call =
			last !== undefined &&
			last.end === p &&
			(last.type === 'function' || last.type === 'call-close' || last.type === 'paren-close');
		this.frames.push(call ? FRAME_CALL : FRAME_GROUP);
		return this.emit(call ? 'call-open' : 'paren-open', p, p + 1);
	}

	/** `#`: the spill operator after a reference or name, otherwise an error literal. */
	private hash(p: number): number {
		const last = this.last;
		if (last !== undefined && last.end === p && (last.type === 'ref' || last.type === 'name')) {
			return this.emit('spill', p, p + 1);
		}
		const code = matchErrorCode(this.src, p);
		if (code === null) {
			return this.unknown(p, p + 1);
		}
		const end = p + code.length;
		return this.emit('error', p, code === '#REF!' ? deadReferenceEnd(this.src, end) : end);
	}
}

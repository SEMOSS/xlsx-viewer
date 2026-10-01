import type { SheetQualifier } from '../ast';
import {
	CH_BANG,
	CH_COLON,
	CH_HASH,
	CH_LBRACKET,
	CH_LPAREN,
	CH_RBRACKET,
	CH_SQUOTE,
	isNameStart,
	isWordChar,
} from './chars';
import {
	isPlainName,
	matchErrorCode,
	parseQualifierText,
	scanA1Area,
	type ScannedArea,
	scanR1C1Area,
	scanWord,
} from './scan';
import type { FormulaTokenType } from './tokens';

/**
 * Lexing of "words" and everything that can carry a sheet prefix:
 * references, names, functions, booleans, structured references and the
 * `'Sheet'!`, `Sheet1:Sheet3!`, `[1]Sheet1!` qualifiers. Each function
 * emits tokens through the lexer and returns the offset after them.
 */
export interface LexerSink {
	readonly src: string;
	/** Base cell for R1C1 references, `undefined` in A1 mode. */
	readonly base: { row: number; col: number } | undefined;
	emit(type: FormulaTokenType, start: number, end: number): number;
	emitQualified(
		type: FormulaTokenType,
		start: number,
		end: number,
		scanned: ScannedArea | null,
		sheet?: SheetQualifier,
		qualifierLength?: number,
	): number;
	unknown(start: number, end: number, incomplete?: boolean): number;
	/** Token from `start` to the end of the text, flagged `incomplete` (unterminated). */
	emitIncomplete(type: FormulaTokenType, start: number): number;
}

/** A word at `s` (starts with `$` or a name character). */
export function lexWord(lx: LexerSink, s: number): number {
	const src = lx.src;
	if (lx.base !== undefined) {
		const r1c1 = scanR1C1Area(src, s, lx.base);
		if (r1c1 !== null) {
			return lx.emitQualified('ref', s, r1c1.end, r1c1);
		}
	}
	const e = scanWord(src, s);
	const next = src.charCodeAt(e);
	const plain = isPlainName(src, s, e);
	if (plain) {
		if (next === CH_BANG) {
			return lexQualified(lx, s, e + 1, { sheet: src.slice(s, e) });
		}
		if (next === CH_COLON) {
			// 3-D reference: `Sheet1:Sheet3!A1`.
			const e2 = scanWord(src, e + 1);
			if (src.charCodeAt(e2) === CH_BANG && isPlainName(src, e + 1, e2)) {
				return lexQualified(lx, s, e2 + 1, {
					sheet: src.slice(s, e),
					sheetEnd: src.slice(e + 1, e2),
				});
			}
		}
		if (next === CH_LPAREN) {
			return lx.emit('function', s, e);
		}
		if (next === CH_LBRACKET) {
			return lexStructured(lx, s, e);
		}
	}
	if (lx.base === undefined) {
		const a1 = scanA1Area(src, s);
		if (a1 !== null) {
			return lx.emitQualified('ref', s, a1.end, a1);
		}
	}
	if (!plain) {
		return lx.unknown(s, e);
	}
	const len = e - s;
	if (len === 4 || len === 5) {
		const upper = src.slice(s, e).toUpperCase();
		if (upper === 'TRUE' || upper === 'FALSE') {
			return lx.emit('boolean', s, e);
		}
	}
	return lx.emit('name', s, e);
}

/** The part after a sheet prefix that starts at `qs`; `p` is just past the `!`. */
function lexQualified(lx: LexerSink, qs: number, p: number, sheet: SheetQualifier): number {
	const src = lx.src;
	const qualifierLength = p - qs;
	const c = src.charCodeAt(p);
	if (c === CH_HASH && matchErrorCode(src, p) === '#REF!') {
		return lx.emitQualified('error', qs, p + 5, null, sheet, qualifierLength);
	}
	const scanned = lx.base !== undefined ? scanR1C1Area(src, p, lx.base) : scanA1Area(src, p);
	if (scanned !== null) {
		return lx.emitQualified('ref', qs, scanned.end, scanned, sheet, qualifierLength);
	}
	if (isNameStart(c)) {
		const e = scanWord(src, p);
		if (isPlainName(src, p, e)) {
			return lx.emitQualified('name', qs, e, null, sheet, qualifierLength);
		}
	}
	// Nothing usable after the `!` (for example `=Sheet1!` while typing).
	return lx.unknown(qs, p, true);
}

/** A quoted sheet prefix: `'My Sheet'!`, `'It''s'!`, `'[Book.xlsx]Data'!`. */
export function lexQuotedQualifier(lx: LexerSink, s: number): number {
	const src = lx.src;
	let i = s + 1;
	while (i < src.length) {
		if (src.charCodeAt(i) === CH_SQUOTE) {
			if (src.charCodeAt(i + 1) !== CH_SQUOTE) {
				break;
			}
			i++;
		}
		i++;
	}
	if (i >= src.length) {
		return lx.unknown(s, src.length, true);
	}
	if (src.charCodeAt(i + 1) !== CH_BANG) {
		return lx.unknown(s, i + 1);
	}
	const raw = src.slice(s + 1, i).replace(/''/g, "'");
	return lexQualified(lx, s, i + 2, parseQualifierText(raw));
}

/**
 * `[` starts an external workbook prefix (`[1]Sheet1!A1`, `[1]!Name`) or a
 * table-less structured reference (`[@Qty]`).
 */
export function lexBracket(lx: LexerSink, s: number): number {
	const src = lx.src;
	let rb = s + 1;
	while (rb < src.length) {
		const c = src.charCodeAt(rb);
		if (c === CH_RBRACKET || c === CH_LBRACKET) {
			break;
		}
		rb++;
	}
	if (rb > s + 1 && src.charCodeAt(rb) === CH_RBRACKET) {
		const workbook = src.slice(s + 1, rb);
		const after = rb + 1;
		if (src.charCodeAt(after) === CH_BANG) {
			return lexQualified(lx, s, after + 1, { sheet: '', workbook });
		}
		const e = scanWord(src, after);
		if (isPlainName(src, after, e)) {
			const n = src.charCodeAt(e);
			if (n === CH_BANG) {
				return lexQualified(lx, s, e + 1, { sheet: src.slice(after, e), workbook });
			}
			if (n === CH_COLON) {
				const e2 = scanWord(src, e + 1);
				if (src.charCodeAt(e2) === CH_BANG && isPlainName(src, e + 1, e2)) {
					const sheet = { sheet: src.slice(after, e), sheetEnd: src.slice(e + 1, e2), workbook };
					return lexQualified(lx, s, e2 + 1, sheet);
				}
			}
		}
	}
	return lexStructured(lx, s, s);
}

/** Structured reference from `s` (table name or `[`) with brackets starting at `b`. */
function lexStructured(lx: LexerSink, s: number, b: number): number {
	const src = lx.src;
	let depth = 0;
	for (let i = b; i < src.length; i++) {
		const c = src.charCodeAt(i);
		if (c === CH_SQUOTE) {
			i++;
		} else if (c === CH_LBRACKET) {
			depth++;
		} else if (c === CH_RBRACKET && --depth === 0) {
			return lx.emit('structured', s, i + 1);
		}
	}
	return lx.emitIncomplete('structured', s);
}

/**
 * End of a dead reference after `#REF!` at `p`: Excel writes references to
 * deleted sheets as `#REF!A1`, `#REF!A1:B2` or `#REF!Name`; they stay one token.
 */
export function deadReferenceEnd(src: string, p: number): number {
	const c = src.charCodeAt(p);
	if (c === CH_HASH) {
		return matchErrorCode(src, p) === '#REF!' ? p + 5 : p;
	}
	if (!isWordChar(c)) {
		return p;
	}
	const area = scanA1Area(src, p);
	if (area !== null) {
		return area.end;
	}
	const e = scanWord(src, p);
	return isPlainName(src, p, e) && src.charCodeAt(e) !== CH_LPAREN ? e : p;
}

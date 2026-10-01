import type { RefCorner, SheetQualifier } from '../ast';

/**
 * Token kinds produced by {@link tokenizeFormula}.
 *
 * Every character of the input belongs to exactly one token, so joining the
 * token texts reproduces the formula exactly (whitespace included).
 */
export type FormulaTokenType =
	/** Leading `=` marking a formula (only at offset 0). */
	| 'equals'
	| 'whitespace'
	/** Whitespace acting as the intersection operator (`A1:B5 B2:C3`). */
	| 'intersect'
	| 'number'
	/** String literal including its quotes; `""` escapes a quote. */
	| 'string'
	| 'boolean'
	/** Error literal, also `Sheet1!#REF!` and dead references such as `#REF!A1`. */
	| 'error'
	/** Cell, area, whole-column or whole-row reference, optionally sheet-qualified. */
	| 'ref'
	/** Defined name, table name or LET / LAMBDA parameter, optionally sheet-qualified. */
	| 'name'
	/** Function name; the `(` that follows is a separate `call-open` token. */
	| 'function'
	/** Structured (table) reference: `Table1[Col]`, `[@Qty]`, `Table1[[#Headers],[A]:[C]]`. */
	| 'structured'
	/** Binary operator: `+ - * / ^ & = <> < > <= >=` and the range operator `:`. */
	| 'operator'
	/** Unary operator: `+`, `-` or the implicit-intersection operator `@`. */
	| 'prefix'
	| 'percent'
	/** Spill-range operator `#` directly after a reference (`A1#`). */
	| 'spill'
	/** `(` opening a group. */
	| 'paren-open'
	| 'paren-close'
	/** `(` opening the argument list of a function or of a called expression. */
	| 'call-open'
	| 'call-close'
	/** `,` separating function arguments. */
	| 'arg-sep'
	/** `,` acting as the union operator (inside a group or at the top level). */
	| 'union'
	| 'array-open'
	| 'array-close'
	| 'array-col-sep'
	| 'array-row-sep'
	/** Anything the grammar does not allow (stray characters, an incomplete sheet prefix). */
	| 'unknown';

export type RefKind = 'cell' | 'cells' | 'columns' | 'rows';

/**
 * A reference area as written: corners keep their order and `$` flags.
 * Whole columns use rows 0 and the last row; whole rows use columns 0 and
 * the last column, so the area always describes real cells.
 */
export interface RefArea {
	kind: RefKind;
	start: RefCorner;
	/** Same object as `start` for single cells. */
	end: RefCorner;
}

export interface FormulaToken {
	type: FormulaTokenType;
	/** Source text of the token. */
	text: string;
	/** Offset of the first character in the formula. */
	start: number;
	/** Offset just past the last character. */
	end: number;
	/** Sheet qualifier of a `ref`, `name` or `error` token. */
	sheet?: SheetQualifier;
	/** Length of the qualifier text including the `!` (`'My Sheet'!` -> 11). */
	qualifierLength?: number;
	/** Parsed reference of a `ref` token. */
	area?: RefArea;
	/** Unterminated string, quoted sheet name or structured reference. */
	incomplete?: boolean;
}

export interface TokenizeOptions {
	/**
	 * Read references in R1C1 notation (`R1C1`, `R[-1]C`, `C2`), resolving
	 * relative parts against this 0-based cell. A1 references are then names.
	 */
	r1c1?: { row: number; col: number };
}

/** Tokens that end an operand: a following `+` / `-` is binary. */
export function endsOperand(type: FormulaTokenType): boolean {
	switch (type) {
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
		case 'ref':
		case 'name':
		case 'structured':
		case 'paren-close':
		case 'call-close':
		case 'array-close':
		case 'percent':
		case 'spill':
			return true;
		default:
			return false;
	}
}

/** Tokens that can end the left operand of an intersection. */
export function endsReferenceOperand(type: FormulaTokenType): boolean {
	switch (type) {
		case 'ref':
		case 'name':
		case 'structured':
		case 'error':
		case 'paren-close':
		case 'call-close':
		case 'spill':
			return true;
		default:
			return false;
	}
}

/** Tokens that can start the right operand of an intersection. */
export function startsReferenceOperand(type: FormulaTokenType): boolean {
	switch (type) {
		case 'ref':
		case 'name':
		case 'structured':
		case 'error':
		case 'function':
		case 'paren-open':
			return true;
		default:
			return false;
	}
}

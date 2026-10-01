import type { ErrorCode } from '../types/cell';

/**
 * Formula abstract syntax tree.
 *
 * Coordinates are 0-based. Relative references keep the absolute row/col of
 * the cell they point to *as written in the formula's own cell*; translating
 * a formula to another cell (copy, fill, shared formulas) is done on the
 * token stream / AST by `translateFormula`.
 */

/** Sheet qualifier of a reference: `Sheet1!A1`, `'My Sheet'!A1`, `Sheet1:Sheet3!A1`, `[1]Sheet1!A1`. */
export interface SheetQualifier {
	/** Sheet name, unquoted. */
	sheet: string;
	/** Last sheet of a 3-D reference (`Sheet1:Sheet3!A1`). */
	sheetEnd?: string;
	/** External workbook reference (`[1]` index or a file name), when present. */
	workbook?: string;
}

/** One corner of a reference. */
export interface RefCorner {
	row: number;
	col: number;
	rowAbs: boolean;
	colAbs: boolean;
}

/** Optional metadata shared by every node; evaluation ignores it. */
export interface AstNodeMeta {
	/**
	 * Number of parenthesis pairs written around this node in the source
	 * (`(A1+B1)*2` puts 1 on the `+` node). Set by the parser so that the
	 * serializer reproduces the author's grouping; absent means none.
	 */
	parens?: number;
}

export interface NumberNode extends AstNodeMeta {
	type: 'number';
	value: number;
}

export interface StringNode extends AstNodeMeta {
	type: 'string';
	value: string;
}

export interface BooleanNode extends AstNodeMeta {
	type: 'boolean';
	value: boolean;
}

export interface ErrorNode extends AstNodeMeta {
	type: 'error';
	code: ErrorCode;
	/** A `#REF!` written in place of a reference may still carry its sheet (`Sheet1!#REF!`). */
	sheet?: SheetQualifier;
}

export interface CellRefNode extends AstNodeMeta {
	type: 'cell';
	sheet?: SheetQualifier;
	ref: RefCorner;
}

/**
 * `A1:B2` (kind `cells`), `A:C` (kind `columns`, rows ignored), `1:3` (kind
 * `rows`, columns ignored). Corners are kept as written (not normalised).
 */
export interface RangeRefNode extends AstNodeMeta {
	type: 'range';
	sheet?: SheetQualifier;
	kind: 'cells' | 'columns' | 'rows';
	start: RefCorner;
	end: RefCorner;
}

/** Defined name, table name or LET / LAMBDA parameter. */
export interface NameNode extends AstNodeMeta {
	type: 'name';
	name: string;
	sheet?: SheetQualifier;
}

export type StructuredSpecifier = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row';

/** Table reference: `Table1[Sales]`, `Table1[[#Headers],[A]:[C]]`, `[@Qty]`. */
export interface StructuredRefNode extends AstNodeMeta {
	type: 'structured';
	/** Table name; omitted for references written inside the table (`[@Qty]`). */
	table?: string;
	specifiers: StructuredSpecifier[];
	columnStart?: string;
	columnEnd?: string;
	/** Source text, used when re-serialising. */
	text: string;
}

export interface FunctionNode extends AstNodeMeta {
	type: 'function';
	/** Canonical upper-case name with file prefixes (`_xlfn.`, `_xlws.`) removed, e.g. `XLOOKUP`. */
	name: string;
	args: AstNode[];
}

/** Calling a non-name expression: `LAMBDA(x, x + 1)(2)`. */
export interface CallNode extends AstNodeMeta {
	type: 'call';
	callee: AstNode;
	args: AstNode[];
}

export interface UnaryNode extends AstNodeMeta {
	type: 'unary';
	/** `@` is the implicit-intersection operator. */
	op: '+' | '-' | '@';
	operand: AstNode;
}

export interface PercentNode extends AstNodeMeta {
	type: 'percent';
	operand: AstNode;
}

export type BinaryOperator =
	| '+'
	| '-'
	| '*'
	| '/'
	| '^'
	| '&'
	| '='
	| '<>'
	| '<'
	| '>'
	| '<='
	| '>='
	/** Range operator between two references (`A1:INDEX(...)`). */
	| ':'
	/** Union of references (`(A1:A3,C1:C3)`). */
	| ','
	/** Intersection of references (`A1:C3 B2:D4`). */
	| ' ';

export interface BinaryNode extends AstNodeMeta {
	type: 'binary';
	op: BinaryOperator;
	left: AstNode;
	right: AstNode;
}

/** Array constant `{1,2;3,4}`: rows of constant elements. */
export interface ArrayNode extends AstNodeMeta {
	type: 'array';
	rows: AstNode[][];
}

/** An omitted argument: `IF(A1,,1)`. */
export interface MissingNode extends AstNodeMeta {
	type: 'missing';
}

/** Spill range reference `A1#`. */
export interface SpillRefNode extends AstNodeMeta {
	type: 'spill';
	ref: CellRefNode;
}

export type AstNode =
	| NumberNode
	| StringNode
	| BooleanNode
	| ErrorNode
	| CellRefNode
	| RangeRefNode
	| NameNode
	| StructuredRefNode
	| FunctionNode
	| CallNode
	| UnaryNode
	| PercentNode
	| BinaryNode
	| ArrayNode
	| MissingNode
	| SpillRefNode;

/**
 * Decoded BIFF8 formula tokens -> formula AST. The token stream is reverse
 * Polish notation, so a stack of AST nodes rebuilds the expression;
 * `serializeFormula` then prints it with Excel's precedence rules (ptgParen
 * becomes the node's `parens` count, ptgAttrSpace is dropped).
 *
 * References: ptgRefN / ptgAreaN hold signed offsets from a base cell (the
 * cell of a shared formula, the top-left cell of a conditional format or
 * validation). In shared formulas and defined names the relative parts of
 * 3-D references are offsets as well; defined names use A1 as their base
 * and wrap negative offsets around the (larger) SpreadsheetML grid, which
 * is how relative names keep their meaning in an `.xlsx`.
 */
import { MAX_COL_INDEX, MAX_COLS, MAX_ROW_INDEX, MAX_ROWS } from '../../../address/constants';
import type {
	AstNode,
	FunctionNode,
	RangeRefNode,
	RefCorner,
	SheetQualifier,
} from '../../../formula/ast';
import { stripFilePrefixes } from '../../../formula/language/names';
import { CETAB } from './cetab';
import { FTAB, UDF_INDEX } from './ftab';
import { type ArrayValue, type Ptg, type PtgLoc, XlsFormulaError } from './ptg-decode';

/** The formula refers to another workbook (or a DDE / OLE source). */
export class XlsExternalReferenceError extends XlsFormulaError {
	constructor(readonly source: string) {
		super(`The formula refers to ${source}`);
		this.name = 'XlsExternalReferenceError';
	}
}

/** An EXTERNSHEET entry: a SUPBOOK and a sheet range in it. */
export interface Xti {
	supbook: number;
	first: number;
	last: number;
}

/** A SUPBOOK record with the EXTERNNAME records that followed it. */
export type SupBook =
	| { kind: 'self'; names: string[] }
	| { kind: 'addin'; names: string[] }
	| { kind: 'external'; path: string; sheets: string[]; names: string[] }
	| { kind: 'other'; names: string[] };

/** Workbook-level lookups a formula needs. */
export interface FormulaEnv {
	/** Sheet names in BOUNDSHEET order (the order EXTERNSHEET sheet indices use). */
	sheetNames: readonly string[];
	xti: readonly Xti[];
	supbooks: readonly SupBook[];
	/** Names of the NAME records in file order (ptgName indices are 1-based). */
	names: readonly string[];
}

/**
 * Where a formula lives: the base cell for relative references and how its
 * references are stored (`cell`: FORMULA / ARRAY, `shared`: SHRFMLA,
 * `name`: NAME, `rule`: conditional formats and data validation).
 */
export interface FormulaPos {
	row: number;
	col: number;
	mode: 'cell' | 'shared' | 'name' | 'rule';
}

const BIFF_ROWS = 65536;
const BIFF_COLS = 256;
const DELETED_SHEET = 0xffff;
const NO_SHEET = 0xfffe;

const int16 = (v: number): number => (v << 16) >> 16;
const int8 = (v: number): number => (v << 24) >> 24;
const wrap = (v: number, size: number): number => ((v % size) + size) % size;

function arrayElement(v: ArrayValue): AstNode {
	if (v === null) {
		return { type: 'string', value: '' };
	}
	switch (typeof v) {
		case 'number':
			return { type: 'number', value: v };
		case 'string':
			return { type: 'string', value: v };
		case 'boolean':
			return { type: 'boolean', value: v };
		default:
			return { type: 'error', code: v.error };
	}
}

class AstBuilder {
	private readonly stack: AstNode[] = [];

	constructor(
		private readonly env: FormulaEnv,
		private readonly pos: FormulaPos,
	) {}

	private pop(): AstNode {
		const node = this.stack.pop();
		if (node === undefined) {
			throw new XlsFormulaError('Malformed formula (missing operand)');
		}
		return node;
	}

	private popArgs(count: number): AstNode[] {
		if (count > this.stack.length) {
			throw new XlsFormulaError('Malformed formula (missing function arguments)');
		}
		return this.stack.splice(this.stack.length - count, count);
	}

	build(ptgs: readonly Ptg[]): AstNode {
		for (const p of ptgs) {
			this.stack.push(this.token(p));
		}
		if (this.stack.length !== 1) {
			throw new XlsFormulaError('Malformed formula (unbalanced operands)');
		}
		return this.stack[0];
	}

	private token(p: Ptg): AstNode {
		switch (p.k) {
			case 'num':
				return { type: 'number', value: p.v };
			case 'str':
				return { type: 'string', value: p.v };
			case 'bool':
				return { type: 'boolean', value: p.v };
			case 'err':
				return { type: 'error', code: p.code };
			case 'miss':
				return { type: 'missing' };
			case 'array':
				return { type: 'array', rows: p.rows.map((r) => r.map(arrayElement)) };
			case 'binary': {
				const right = this.pop();
				return { type: 'binary', op: p.op, left: this.pop(), right };
			}
			case 'unary':
				return { type: 'unary', op: p.op, operand: this.pop() };
			case 'percent':
				return { type: 'percent', operand: this.pop() };
			case 'paren': {
				const node = this.pop();
				node.parens = (node.parens ?? 0) + 1;
				return node;
			}
			case 'sum':
				return { type: 'function', name: 'SUM', args: [this.pop()] };
			case 'func':
				return this.call(p.index, p.argc, p.fixed, p.command);
			case 'name':
				return this.name(p.index);
			case 'nameX':
				return this.nameX(p.ixti, p.index);
			case 'ref':
				return this.ref(p.loc, p.n, p.ixti);
			case 'area':
				return this.area(p.first, p.last, p.n, p.ixti);
			case 'refErr': {
				const sheet = p.ixti === undefined ? undefined : this.qualifier(p.ixti);
				return sheet ? { type: 'error', code: '#REF!', sheet } : { type: 'error', code: '#REF!' };
			}
			case 'exp':
			case 'tbl':
				throw new XlsFormulaError('Shared, array or table reference inside a formula');
		}
	}

	private call(index: number, argc: number, fixed: boolean, command: boolean): FunctionNode {
		if (command) {
			const name = CETAB.get(index);
			if (!name) {
				throw new XlsFormulaError(`Unknown macro command ${index}`);
			}
			return { type: 'function', name, args: this.popArgs(argc) };
		}
		if (index === UDF_INDEX && !fixed) {
			const args = this.popArgs(argc);
			const head = args.shift();
			if (!head || head.type !== 'name') {
				throw new XlsFormulaError('A user-defined function call has no function name');
			}
			return { type: 'function', name: stripFilePrefixes(head.name), args };
		}
		const fn = FTAB.get(index);
		if (!fn) {
			throw new XlsFormulaError(`Unknown function index ${index}`);
		}
		return { type: 'function', name: fn.name, args: this.popArgs(fixed ? fn.min : argc) };
	}

	private name(index: number): AstNode {
		const name = this.env.names[index - 1];
		return name === undefined ? { type: 'error', code: '#NAME?' } : { type: 'name', name };
	}

	private nameX(ixti: number, index: number): AstNode {
		const x = this.env.xti[ixti];
		const book = x ? this.env.supbooks[x.supbook] : undefined;
		if (!x || !book) {
			throw new XlsFormulaError(`Missing EXTERNSHEET entry ${ixti}`);
		}
		if (book.kind === 'self') {
			const name = this.env.names[index - 1];
			if (name === undefined) {
				return { type: 'error', code: '#NAME?' };
			}
			const sheet = x.first === NO_SHEET ? undefined : this.env.sheetNames[x.first];
			return sheet === undefined
				? { type: 'name', name }
				: { type: 'name', name, sheet: { sheet } };
		}
		if (book.kind === 'addin') {
			const name = book.names[index - 1];
			if (!name) {
				throw new XlsFormulaError(`Missing add-in function name ${index}`);
			}
			return { type: 'name', name };
		}
		throw new XlsExternalReferenceError(
			book.kind === 'external' ? book.path : 'an external source',
		);
	}

	/** Sheet qualifier of an EXTERNSHEET entry: `null` for a deleted sheet, `undefined` for none. */
	private qualifier(ixti: number): SheetQualifier | null | undefined {
		const x = this.env.xti[ixti];
		const book = x ? this.env.supbooks[x.supbook] : undefined;
		if (!x || !book) {
			throw new XlsFormulaError(`Missing EXTERNSHEET entry ${ixti}`);
		}
		if (book.kind !== 'self') {
			throw new XlsExternalReferenceError(
				book.kind === 'external' ? book.path : 'an external source',
			);
		}
		if (x.first === NO_SHEET) {
			return undefined;
		}
		const first = this.env.sheetNames[x.first];
		if (x.first === DELETED_SHEET || first === undefined) {
			return null;
		}
		const last = this.env.sheetNames[x.last];
		return last !== undefined && x.last !== x.first
			? { sheet: first, sheetEnd: last }
			: { sheet: first };
	}

	/** Relative parts are offsets for N tokens, in names, and for 3-D tokens in shared formulas. */
	private offsets(n: boolean, ixti: number | undefined): boolean {
		const mode = this.pos.mode;
		return n || mode === 'name' || (ixti !== undefined && mode === 'shared');
	}

	private corner(l: PtgLoc, offsets: boolean): RefCorner {
		const rowRel = (l.col & 0x8000) !== 0;
		const colRel = (l.col & 0x4000) !== 0;
		const named = this.pos.mode === 'name';
		const rows = named ? MAX_ROWS : BIFF_ROWS;
		const cols = named ? MAX_COLS : BIFF_COLS;
		const row = offsets && rowRel ? wrap(this.pos.row + int16(l.rw), rows) : l.rw;
		const col =
			offsets && colRel
				? wrap(this.pos.col + int8(l.col & 0xff), cols)
				: Math.min(l.col & 0x3fff, MAX_COL_INDEX);
		return { row, col, rowAbs: !rowRel, colAbs: !colRel };
	}

	private ref(l: PtgLoc, n: boolean, ixti: number | undefined): AstNode {
		const sheet = ixti === undefined ? undefined : this.qualifier(ixti);
		if (sheet === null) {
			return { type: 'error', code: '#REF!' };
		}
		const ref = this.corner(l, this.offsets(n, ixti));
		return sheet ? { type: 'cell', ref, sheet } : { type: 'cell', ref };
	}

	private area(first: PtgLoc, last: PtgLoc, n: boolean, ixti: number | undefined): AstNode {
		const sheet = ixti === undefined ? undefined : this.qualifier(ixti);
		if (sheet === null) {
			return { type: 'error', code: '#REF!' };
		}
		const offsets = this.offsets(n, ixti);
		const a = this.corner(first, offsets);
		const b = this.corner(last, offsets);
		const allRows = a.row === 0 && (b.row === BIFF_ROWS - 1 || b.row === MAX_ROW_INDEX);
		const allCols = a.col === 0 && (b.col === BIFF_COLS - 1 || b.col === MAX_COL_INDEX);
		let node: RangeRefNode;
		if (allRows && !allCols) {
			node = {
				type: 'range',
				kind: 'columns',
				start: { row: 0, col: a.col, rowAbs: false, colAbs: a.colAbs },
				end: { row: MAX_ROW_INDEX, col: b.col, rowAbs: false, colAbs: b.colAbs },
			};
		} else if (allCols) {
			node = {
				type: 'range',
				kind: 'rows',
				start: { row: a.row, col: 0, rowAbs: a.rowAbs, colAbs: false },
				end: { row: b.row, col: MAX_COL_INDEX, rowAbs: b.rowAbs, colAbs: false },
			};
		} else {
			node = { type: 'range', kind: 'cells', start: a, end: b };
		}
		if (sheet) {
			node.sheet = sheet;
		}
		return node;
	}
}

/** Build the AST of a decoded formula. Throws {@link XlsFormulaError}. */
export function ptgsToAst(ptgs: readonly Ptg[], env: FormulaEnv, pos: FormulaPos): AstNode {
	return new AstBuilder(env, pos).build(ptgs);
}

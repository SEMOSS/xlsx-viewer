/**
 * BIFF8 parsed formulas ([MS-XLS] 2.5.198): an `Rgce` token stream in
 * reverse Polish notation plus `RgbExtra` data (array constants, cached
 * areas of ptgMemArea). This module only decodes the binary layout into
 * {@link Ptg} tokens; `ptg-ast.ts` gives them meaning.
 */
import type { BinaryOperator } from '../../../formula/ast';
import type { ErrorCode } from '../../../types/cell';
import { BiffCursor, BiffTruncatedError } from '../biff-cursor';

/** A formula could not be decoded or rendered (the cell keeps its cached value). */
export class XlsFormulaError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'XlsFormulaError';
	}
}

/** Element of an array constant (`null` is an empty element). */
export type ArrayValue = number | string | boolean | { error: ErrorCode } | null;

/** Cell reference as stored: row and the column word with its relative flags. */
export interface PtgLoc {
	rw: number;
	/** Column (low 14 bits) with fColRelative (0x4000) and fRwRelative (0x8000). */
	col: number;
}

export type Ptg =
	| { k: 'num'; v: number }
	| { k: 'str'; v: string }
	| { k: 'bool'; v: boolean }
	| { k: 'err'; code: ErrorCode }
	| { k: 'miss' }
	| { k: 'array'; rows: ArrayValue[][] }
	| { k: 'binary'; op: BinaryOperator }
	| { k: 'unary'; op: '+' | '-' }
	| { k: 'percent' }
	| { k: 'paren' }
	/** ptgFunc (`fixed`) or ptgFuncVar; `command` selects the Cetab (macro command) table. */
	| { k: 'func'; index: number; argc: number; fixed: boolean; command: boolean }
	/** ptgAttrSum: SUM of the single operand on the stack. */
	| { k: 'sum' }
	| { k: 'name'; index: number }
	| { k: 'nameX'; ixti: number; index: number }
	/** ptgRef / ptgRefN (`n`) / ptgRef3d (`ixti`). */
	| { k: 'ref'; loc: PtgLoc; n: boolean; ixti?: number }
	/** ptgArea / ptgAreaN (`n`) / ptgArea3d (`ixti`). */
	| { k: 'area'; first: PtgLoc; last: PtgLoc; n: boolean; ixti?: number }
	/** ptgRefErr / ptgAreaErr and their 3-D forms. */
	| { k: 'refErr'; ixti?: number }
	| { k: 'exp'; row: number; col: number }
	| { k: 'tbl'; row: number; col: number };

/** BErr error codes. */
export const ERROR_CODES: Readonly<Record<number, ErrorCode>> = {
	0x00: '#NULL!',
	0x07: '#DIV/0!',
	0x0f: '#VALUE!',
	0x17: '#REF!',
	0x1d: '#NAME?',
	0x24: '#NUM!',
	0x2a: '#N/A',
	0x2b: '#GETTING_DATA',
};

const BINARY: Readonly<Record<number, BinaryOperator>> = {
	0x03: '+',
	0x04: '-',
	0x05: '*',
	0x06: '/',
	0x07: '^',
	0x08: '&',
	0x09: '<',
	0x0a: '<=',
	0x0b: '=',
	0x0c: '>=',
	0x0d: '>',
	0x0e: '<>',
	0x0f: ' ',
	0x10: ',',
	0x11: ':',
};

/** Tokens per formula; BIFF8 formulas hold at most a few thousand bytes. */
const MAX_TOKENS = 10_000;

export function errorCode(code: number): ErrorCode {
	const e = ERROR_CODES[code];
	if (!e) {
		throw new XlsFormulaError(`Unknown error code 0x${code.toString(16)}`);
	}
	return e;
}

function loc(cur: BiffCursor): PtgLoc {
	return { rw: cur.u16(), col: cur.u16() };
}

/** PtgExtraArray: dimensions and `SerAr` elements, row by row. */
function readArray(ext: BiffCursor): ArrayValue[][] {
	const cols = ext.u8() + 1;
	const rows = ext.u16() + 1;
	if (rows * cols * 4 > ext.remaining) {
		throw new XlsFormulaError('The array constant is truncated');
	}
	const out: ArrayValue[][] = [];
	for (let r = 0; r < rows; r++) {
		const row: ArrayValue[] = [];
		for (let c = 0; c < cols; c++) {
			const type = ext.u8();
			switch (type) {
				case 0x00:
					ext.skip(8);
					row.push(null);
					break;
				case 0x01:
					row.push(ext.f64());
					break;
				case 0x02:
					row.push(ext.unicodeString());
					break;
				case 0x04:
					row.push(ext.u8() !== 0);
					ext.skip(7);
					break;
				case 0x10:
					row.push({ error: errorCode(ext.u8()) });
					ext.skip(7);
					break;
				default:
					throw new XlsFormulaError(`Unknown array element type ${type}`);
			}
		}
		out.push(row);
	}
	return out;
}

/** ptgAttr: only SUM shortcuts produce a token; spaces, volatility and jumps are skipped. */
function readAttr(cur: BiffCursor, out: Ptg[]): void {
	const flags = cur.u8();
	if (flags & 0x04) {
		const count = cur.u16();
		cur.skip(2 * (count + 1));
	} else {
		cur.skip(2);
		if (flags & 0x10) {
			out.push({ k: 'sum' });
		}
	}
}

/** Control and constant tokens (ptg < 0x20). */
function readBase(ptg: number, cur: BiffCursor, out: Ptg[]): void {
	const op = BINARY[ptg];
	if (op) {
		out.push({ k: 'binary', op });
		return;
	}
	switch (ptg) {
		case 0x01:
		case 0x02: {
			const row = cur.u16();
			const col = cur.u16();
			out.push(ptg === 0x01 ? { k: 'exp', row, col } : { k: 'tbl', row, col });
			return;
		}
		case 0x12:
		case 0x13:
			out.push({ k: 'unary', op: ptg === 0x12 ? '+' : '-' });
			return;
		case 0x14:
			out.push({ k: 'percent' });
			return;
		case 0x15:
			out.push({ k: 'paren' });
			return;
		case 0x16:
			out.push({ k: 'miss' });
			return;
		case 0x17:
			out.push({ k: 'str', v: cur.shortString() });
			return;
		case 0x19:
			readAttr(cur, out);
			return;
		case 0x1c:
			out.push({ k: 'err', code: errorCode(cur.u8()) });
			return;
		case 0x1d:
			out.push({ k: 'bool', v: cur.u8() !== 0 });
			return;
		case 0x1e:
			out.push({ k: 'num', v: cur.u16() });
			return;
		case 0x1f:
			out.push({ k: 'num', v: cur.f64() });
			return;
		case 0x18:
			throw new XlsFormulaError('Extended formula tokens (ptgElf, ptgSxName) are not supported');
		default:
			throw new XlsFormulaError(`Unknown formula token 0x${ptg.toString(16)}`);
	}
}

/** Operand tokens (ptg >= 0x20): the low five bits select the token, bits 5-6 the class. */
function readOperand(ptg: number, cur: BiffCursor, ext: BiffCursor, out: Ptg[]): void {
	switch (ptg & 0x1f) {
		case 0x00:
			cur.skip(7);
			out.push({ k: 'array', rows: readArray(ext) });
			return;
		case 0x01:
			out.push({ k: 'func', index: cur.u16(), argc: 0, fixed: true, command: false });
			return;
		case 0x02: {
			const argc = cur.u8() & 0x7f;
			const tab = cur.u16();
			out.push({
				k: 'func',
				index: tab & 0x7fff,
				argc,
				fixed: false,
				command: (tab & 0x8000) !== 0,
			});
			return;
		}
		case 0x03:
			out.push({ k: 'name', index: cur.u16() });
			cur.skip(2);
			return;
		case 0x04:
		case 0x0c:
			out.push({ k: 'ref', loc: loc(cur), n: (ptg & 0x1f) === 0x0c });
			return;
		case 0x05:
		case 0x0d:
			out.push({ k: 'area', ...areaLocs(cur), n: (ptg & 0x1f) === 0x0d });
			return;
		case 0x06:
			// ptgMemArea: the subexpression follows inline; its cached areas sit in the extra data
			// (some writers leave them out).
			cur.skip(6);
			if (ext.remaining >= 2) {
				ext.skip(Math.min(ext.u16() * 8, ext.remaining));
			}
			return;
		case 0x07:
		case 0x08:
			cur.skip(6);
			return;
		case 0x09:
		case 0x0e:
		case 0x0f:
			cur.skip(2);
			return;
		case 0x0a:
			cur.skip(4);
			out.push({ k: 'refErr' });
			return;
		case 0x0b:
			cur.skip(8);
			out.push({ k: 'refErr' });
			return;
		case 0x19: {
			const ixti = cur.u16();
			out.push({ k: 'nameX', ixti, index: cur.u16() });
			cur.skip(2);
			return;
		}
		case 0x1a: {
			const ixti = cur.u16();
			out.push({ k: 'ref', loc: loc(cur), n: false, ixti });
			return;
		}
		case 0x1b: {
			const ixti = cur.u16();
			out.push({ k: 'area', ...areaLocs(cur), n: false, ixti });
			return;
		}
		case 0x1c:
		case 0x1d: {
			const ixti = cur.u16();
			cur.skip((ptg & 0x1f) === 0x1c ? 4 : 8);
			out.push({ k: 'refErr', ixti });
			return;
		}
		default:
			throw new XlsFormulaError(`Unknown formula token 0x${ptg.toString(16)}`);
	}
}

function areaLocs(cur: BiffCursor): { first: PtgLoc; last: PtgLoc } {
	const rw1 = cur.u16();
	const rw2 = cur.u16();
	const col1 = cur.u16();
	const col2 = cur.u16();
	return { first: { rw: rw1, col: col1 }, last: { rw: rw2, col: col2 } };
}

/**
 * Decode an `Rgce` (and the `RgbExtra` bytes that follow it in the
 * record). Throws {@link XlsFormulaError} for unknown or truncated tokens.
 */
export function decodeRgce(rgce: Uint8Array, extra: Uint8Array = new Uint8Array(0)): Ptg[] {
	const cur = new BiffCursor(rgce);
	const ext = new BiffCursor(extra);
	const out: Ptg[] = [];
	try {
		while (cur.remaining > 0) {
			if (out.length > MAX_TOKENS) {
				throw new XlsFormulaError('The formula is too long');
			}
			const ptg = cur.u8();
			if (ptg < 0x20) {
				readBase(ptg, cur, out);
			} else if (ptg < 0x80) {
				readOperand(ptg, cur, ext, out);
			} else {
				throw new XlsFormulaError(`Unknown formula token 0x${ptg.toString(16)}`);
			}
		}
	} catch (e) {
		if (e instanceof BiffTruncatedError) {
			throw new XlsFormulaError('The formula is truncated');
		}
		throw e;
	}
	return out;
}

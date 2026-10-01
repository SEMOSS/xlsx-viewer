/**
 * Test-only builders for BIFF8 formula tokens. Each returns the token bytes;
 * concatenate them in reverse Polish order. Reference tokens default to the
 * reference class (0x20); pass `cls: 0x40` (value) or `0x60` (array).
 */
import { f64, shortString, u16, u32, xlString } from './bytes';

export interface RefFlags {
	rowAbs?: boolean;
	colAbs?: boolean;
	cls?: 0x20 | 0x40 | 0x60;
}

function colWord(col: number, f: RefFlags): number {
	return (col & 0x3fff) | (f.colAbs ? 0 : 0x4000) | (f.rowAbs ? 0 : 0x8000);
}

/** Two's complement 8-bit column offset / 16-bit row offset for ptgRefN-style tokens. */
const off8 = (v: number): number => v & 0xff;
const off16 = (v: number): number => v & 0xffff;

export const ptg = {
	exp: (row: number, col: number): number[] => [0x01, ...u16(row), ...u16(col)],
	tbl: (row: number, col: number): number[] => [0x02, ...u16(row), ...u16(col)],
	add: [0x03],
	sub: [0x04],
	mul: [0x05],
	div: [0x06],
	power: [0x07],
	concat: [0x08],
	lt: [0x09],
	le: [0x0a],
	eq: [0x0b],
	ge: [0x0c],
	gt: [0x0d],
	ne: [0x0e],
	isect: [0x0f],
	union: [0x10],
	range: [0x11],
	uplus: [0x12],
	uminus: [0x13],
	percent: [0x14],
	paren: [0x15],
	miss: [0x16],
	str: (s: string, high?: boolean): number[] => [0x17, ...shortString(s, high)],
	err: (code: number): number[] => [0x1c, code],
	bool: (v: boolean): number[] => [0x1d, v ? 1 : 0],
	int: (v: number): number[] => [0x1e, ...u16(v)],
	num: (v: number): number[] => [0x1f, ...f64(v)],
	attrSemi: [0x19, 0x01, 0, 0],
	attrIf: (offset: number): number[] => [0x19, 0x02, ...u16(offset)],
	attrChoose: (offsets: number[]): number[] => [
		0x19,
		0x04,
		...u16(offsets.length - 1),
		...offsets.flatMap(u16),
	],
	attrGoto: (offset: number): number[] => [0x19, 0x08, ...u16(offset)],
	attrSum: [0x19, 0x10, 0, 0],
	attrSpace: (count: number, type = 0): number[] => [0x19, 0x40, type, count],
	/** ptgArray (data goes to the extra bytes, see {@link arrayExtra}). */
	array: (cls = 0x40): number[] => [cls, 0, 0, 0, 0, 0, 0, 0],
	func: (index: number, cls = 0x40): number[] => [cls | 0x01, ...u16(index)],
	funcVar: (
		index: number,
		argc: number,
		opts: { command?: boolean; cls?: number } = {},
	): number[] => [(opts.cls ?? 0x40) | 0x02, argc, ...u16(index | (opts.command ? 0x8000 : 0))],
	name: (index: number, cls = 0x20): number[] => [cls | 0x03, ...u32(index)],
	nameX: (ixti: number, index: number, cls = 0x20): number[] => [
		cls | 0x19,
		...u16(ixti),
		...u32(index),
	],
	ref: (row: number, col: number, f: RefFlags = {}): number[] => [
		(f.cls ?? 0x20) | 0x04,
		...u16(row),
		...u16(colWord(col, f)),
	],
	area: (r1: number, c1: number, r2: number, c2: number, f: RefFlags = {}): number[] => [
		(f.cls ?? 0x20) | 0x05,
		...u16(r1),
		...u16(r2),
		...u16(colWord(c1, f)),
		...u16(colWord(c2, f)),
	],
	/** ptgRefN with row / column offsets (relative parts) or absolute values. */
	refN: (row: number, col: number, f: RefFlags = {}): number[] => [
		(f.cls ?? 0x20) | 0x0c,
		...u16(f.rowAbs ? row : off16(row)),
		...u16(colWord(f.colAbs ? col : off8(col), f)),
	],
	areaN: (r1: number, c1: number, r2: number, c2: number, f: RefFlags = {}): number[] => [
		(f.cls ?? 0x20) | 0x0d,
		...u16(f.rowAbs ? r1 : off16(r1)),
		...u16(f.rowAbs ? r2 : off16(r2)),
		...u16(colWord(f.colAbs ? c1 : off8(c1), f)),
		...u16(colWord(f.colAbs ? c2 : off8(c2), f)),
	],
	refErr: (cls = 0x20): number[] => [cls | 0x0a, 0, 0, 0, 0],
	areaErr: (cls = 0x20): number[] => [cls | 0x0b, 0, 0, 0, 0, 0, 0, 0, 0],
	ref3d: (ixti: number, row: number, col: number, f: RefFlags = {}): number[] => [
		(f.cls ?? 0x20) | 0x1a,
		...u16(ixti),
		...u16(row),
		...u16(colWord(col, f)),
	],
	area3d: (
		ixti: number,
		r1: number,
		c1: number,
		r2: number,
		c2: number,
		f: RefFlags = {},
	): number[] => [
		(f.cls ?? 0x20) | 0x1b,
		...u16(ixti),
		...u16(r1),
		...u16(r2),
		...u16(colWord(c1, f)),
		...u16(colWord(c2, f)),
	],
	refErr3d: (ixti: number, cls = 0x20): number[] => [cls | 0x1c, ...u16(ixti), 0, 0, 0, 0],
	areaErr3d: (ixti: number, cls = 0x20): number[] => [
		cls | 0x1d,
		...u16(ixti),
		...[0, 0, 0, 0, 0, 0, 0, 0],
	],
	memArea: (cce: number, cls = 0x20): number[] => [cls | 0x06, 0, 0, 0, 0, ...u16(cce)],
	memErr: (cce: number, cls = 0x20): number[] => [cls | 0x07, 0, 0, 0, 0, ...u16(cce)],
	memNoMem: (cce: number, cls = 0x20): number[] => [cls | 0x08, 0, 0, 0, 0, ...u16(cce)],
	memFunc: (cce: number, cls = 0x20): number[] => [cls | 0x09, ...u16(cce)],
};

export type ArrayItem = number | string | boolean | { error: number } | null;

/** PtgExtraArray: columns - 1, rows - 1, then `SerAr` elements row by row. */
export function arrayExtra(rows: ArrayItem[][]): number[] {
	const out = [rows[0].length - 1, ...u16(rows.length - 1)];
	for (const row of rows) {
		for (const v of row) {
			if (v === null) {
				out.push(0x00, 0, 0, 0, 0, 0, 0, 0, 0);
			} else if (typeof v === 'number') {
				out.push(0x01, ...f64(v));
			} else if (typeof v === 'string') {
				out.push(0x02, ...xlString(v));
			} else if (typeof v === 'boolean') {
				out.push(0x04, v ? 1 : 0, 0, 0, 0, 0, 0, 0, 0);
			} else {
				out.push(0x10, v.error, 0, 0, 0, 0, 0, 0, 0);
			}
		}
	}
	return out;
}

/** PtgExtraMem: cached areas of a ptgMemArea. */
export function memExtra(areas: Array<[number, number, number, number]>): number[] {
	return [
		...u16(areas.length),
		...areas.flatMap(([r1, r2, c1, c2]) => [...u16(r1), ...u16(r2), ...u16(c1), ...u16(c2)]),
	];
}

/** Flatten token byte arrays. */
export function rgce(...tokens: number[][]): Uint8Array {
	return Uint8Array.from(tokens.flat());
}

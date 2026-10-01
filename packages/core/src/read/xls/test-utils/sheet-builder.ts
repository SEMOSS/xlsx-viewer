/** Test-only builders for BIFF8 worksheet records (links, notes and rules: `feature-builder.ts`). */
import { RT } from '../biff-opcodes';
import { f64, record, ref8, u16, u32, xlString } from './bytes';

const cell = (row: number, col: number, xf: number): number[] => [
	...u16(row),
	...u16(col),
	...u16(xf),
];
const refU = (r1: number, r2: number, c1: number, c2: number): number[] => [
	...u16(r1),
	...u16(r2),
	c1,
	c2,
];

export const number = (row: number, col: number, xf: number, v: number): number[] =>
	record(RT.NUMBER, [...cell(row, col, xf), ...f64(v)]);

/** RK from raw bits (see {@link rkInt}, {@link rkFloat}). */
export const rk = (row: number, col: number, xf: number, bits: number): number[] =>
	record(RT.RK, [...cell(row, col, xf), ...u32(bits)]);

/** RK bits of an integer (optionally scaled by 100). */
export const rkInt = (v: number, x100 = false): number => ((v << 2) | 0x02 | (x100 ? 1 : 0)) >>> 0;

/** RK bits of a double whose low 34 bits are zero (optionally scaled by 100). */
export function rkFloat(v: number, x100 = false): number {
	const b = new DataView(new ArrayBuffer(8));
	b.setFloat64(0, v, true);
	return ((b.getUint32(4, true) & 0xfffffffc) | (x100 ? 1 : 0)) >>> 0;
}

export const mulrk = (row: number, first: number, items: Array<[number, number]>): number[] =>
	record(RT.MULRK, [
		...u16(row),
		...u16(first),
		...items.flatMap(([xf, bits]) => [...u16(xf), ...u32(bits)]),
		...u16(first + items.length - 1),
	]);

export const labelSst = (row: number, col: number, xf: number, index: number): number[] =>
	record(RT.LABELSST, [...cell(row, col, xf), ...u32(index)]);

export const label = (row: number, col: number, xf: number, text: string): number[] =>
	record(RT.LABEL, [...cell(row, col, xf), ...xlString(text)]);

export const rstring = (
	row: number,
	col: number,
	xf: number,
	text: string,
	runs: Array<[number, number]>,
): number[] =>
	record(RT.RSTRING, [
		...cell(row, col, xf),
		...xlString(text),
		...u16(runs.length),
		...runs.flatMap(([i, f]) => [...u16(i), ...u16(f)]),
	]);

export const boolErr = (
	row: number,
	col: number,
	xf: number,
	v: number,
	isError: boolean,
): number[] => record(RT.BOOLERR, [...cell(row, col, xf), v, isError ? 1 : 0]);

export const blank = (row: number, col: number, xf: number): number[] =>
	record(RT.BLANK, cell(row, col, xf));

export const mulblank = (row: number, first: number, xfs: number[]): number[] =>
	record(RT.MULBLANK, [
		...u16(row),
		...u16(first),
		...xfs.flatMap(u16),
		...u16(first + xfs.length - 1),
	]);

export type FormulaResult = number | string | boolean | { error: number } | null;

function resultBytes(v: FormulaResult): number[] {
	if (typeof v === 'number') {
		return f64(v);
	}
	if (typeof v === 'string') {
		return v === '' ? [3, 0, 0, 0, 0, 0, 0xff, 0xff] : [0, 0, 0, 0, 0, 0, 0xff, 0xff];
	}
	if (typeof v === 'boolean') {
		return [1, 0, v ? 1 : 0, 0, 0, 0, 0xff, 0xff];
	}
	if (v === null) {
		return [3, 0, 0, 0, 0, 0, 0xff, 0xff];
	}
	return [2, 0, v.error, 0, 0, 0, 0xff, 0xff];
}

/** FORMULA record; a non-empty string result needs a following {@link stringResult}. */
export function formula(
	row: number,
	col: number,
	xf: number,
	result: FormulaResult,
	rgce: ArrayLike<number>,
	extra: number[] = [],
	flags = 0,
): number[] {
	return record(RT.FORMULA, [
		...cell(row, col, xf),
		...resultBytes(result),
		...u16(flags),
		...u32(0),
		...u16(rgce.length),
		...Array.from(rgce),
		...extra,
	]);
}

export const stringResult = (text: string): number[] => record(RT.STRING, xlString(text));

export const shrfmla = (
	r1: number,
	r2: number,
	c1: number,
	c2: number,
	rgce: ArrayLike<number>,
	extra: number[] = [],
): number[] =>
	record(RT.SHRFMLA, [
		...refU(r1, r2, c1, c2),
		0,
		(r2 - r1 + 1) * (c2 - c1 + 1),
		...u16(rgce.length),
		...Array.from(rgce),
		...extra,
	]);

export const arrayFormula = (
	r1: number,
	r2: number,
	c1: number,
	c2: number,
	rgce: ArrayLike<number>,
	extra: number[] = [],
): number[] =>
	record(RT.ARRAY, [
		...refU(r1, r2, c1, c2),
		...u16(0),
		...u32(0),
		...u16(rgce.length),
		...Array.from(rgce),
		...extra,
	]);

export const table = (r1: number, r2: number, c1: number, c2: number): number[] =>
	record(RT.TABLE, [
		...refU(r1, r2, c1, c2),
		...u16(0),
		...u16(0),
		...u16(0),
		...u16(0),
		...u16(0),
	]);

export interface RowOptions {
	height?: number;
	custom?: boolean;
	hidden?: boolean;
	level?: number;
	collapsed?: boolean;
	xf?: number;
}

export function row(r: number, o: RowOptions = {}): number[] {
	const flags =
		(o.level ?? 0) |
		(o.collapsed ? 0x10 : 0) |
		(o.hidden ? 0x20 : 0) |
		(o.custom ? 0x40 : 0) |
		(o.xf !== undefined ? 0x80 : 0) |
		0x100;
	return record(RT.ROW, [
		...u16(r),
		...u16(0),
		...u16(1),
		...u16(o.height ?? 255),
		...u16(0),
		...u16(0),
		...u16(flags),
		...u16(o.xf ?? 15),
	]);
}

export function colInfo(
	first: number,
	last: number,
	width: number,
	o: { xf?: number; hidden?: boolean; custom?: boolean; level?: number } = {},
): number[] {
	const flags = (o.hidden ? 0x01 : 0) | (o.custom ? 0x02 : 0) | ((o.level ?? 0) << 8);
	return record(RT.COLINFO, [
		...u16(first),
		...u16(last),
		...u16(width),
		...u16(o.xf ?? 15),
		...u16(flags),
		...u16(0),
	]);
}

export const defColWidth = (chars: number): number[] => record(RT.DEFCOLWIDTH, u16(chars));
export const defaultRowHeight = (twips: number, flags = 0): number[] =>
	record(RT.DEFAULTROWHEIGHT, [...u16(flags), ...u16(twips)]);
export const mergedCells = (ranges: Array<[number, number, number, number]>): number[] =>
	record(RT.MERGEDCELLS, [
		...u16(ranges.length),
		...ranges.flatMap(([r1, r2, c1, c2]) => ref8(r1, r2, c1, c2)),
	]);

export function window2(
	o: { flags?: number; top?: number; left?: number; grid?: number; zoom?: number } = {},
): number[] {
	return record(RT.WINDOW2, [
		...u16(o.flags ?? 0x06b6),
		...u16(o.top ?? 0),
		...u16(o.left ?? 0),
		...u16(o.grid ?? 64),
		...u16(0),
		...u16(0),
		...u16(o.zoom ?? 0),
		...u32(0),
	]);
}

export const scl = (num: number, den: number): number[] =>
	record(RT.SCL, [...u16(num), ...u16(den)]);
export const pane = (x: number, y: number, top: number, left: number, active: number): number[] =>
	record(RT.PANE, [...u16(x), ...u16(y), ...u16(top), ...u16(left), active, 0]);
export const selection = (
	paneId: number,
	r: number,
	c: number,
	ranges: Array<[number, number, number, number]>,
): number[] =>
	record(RT.SELECTION, [
		paneId,
		...u16(r),
		...u16(c),
		...u16(0),
		...u16(ranges.length),
		...ranges.flatMap(([r1, r2, c1, c2]) => refU(r1, r2, c1, c2)),
	]);

export const wsbool = (flags: number): number[] => record(RT.WSBOOL, u16(flags));
export const protect = (on = true): number[] => record(RT.PROTECT, u16(on ? 1 : 0));
export const margin = (type: number, inches: number): number[] => record(type, f64(inches));
export const header = (text: string): number[] => record(RT.HEADER, xlString(text));
export const calcMode = (mode: number): number[] => record(RT.CALCMODE, u16(mode));

export * from './feature-builder';

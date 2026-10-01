/**
 * Test-only BIFF8 workbook builder: record helpers for the workbook globals
 * and the assembly of a `Workbook` stream (globals, BOUNDSHEET offsets,
 * sheet substreams) wrapped in a compound file. Sheet records live in
 * `sheet-builder.ts`.
 */
import { BOF_TYPE, RT } from '../biff-opcodes';
import { concat, record, shortString, stringNoCch, u16, u32, xlString } from './bytes';
import { type CfbStream, writeCompoundFileForTest } from './cfb-writer';

export function bof(type: number = BOF_TYPE.worksheet, version = 0x0600): number[] {
	return record(RT.BOF, [
		...u16(version),
		...u16(type),
		...u16(0x0dbb),
		...u16(0x07cc),
		...u32(0x41),
		...u32(0x06),
	]);
}

export const eof = (): number[] => record(RT.EOF, []);

/** A record followed by CONTINUE records, one per extra chunk. */
export function continued(type: number, chunks: number[][]): number[] {
	return [record(type, chunks[0]), ...chunks.slice(1).map((c) => record(RT.CONTINUE, c))].flat();
}

export const dateMode = (is1904: boolean): number[] => record(RT.DATEMODE, u16(is1904 ? 1 : 0));

export function window1(
	opts: { activeTab?: number; firstTab?: number; flags?: number; tabRatio?: number } = {},
): number[] {
	return record(RT.WINDOW1, [
		...u16(0),
		...u16(0),
		...u16(0x3000),
		...u16(0x2000),
		...u16(opts.flags ?? 0x38),
		...u16(opts.activeTab ?? 0),
		...u16(opts.firstTab ?? 0),
		...u16(1),
		...u16(opts.tabRatio ?? 600),
	]);
}

export interface FontOptions {
	name?: string;
	/** Height in twips (1/20 pt). */
	height?: number;
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	underline?: number;
	color?: number;
	script?: number;
	family?: number;
	charset?: number;
}

export function font(o: FontOptions = {}): number[] {
	const flags = (o.italic ? 0x02 : 0) | (o.strike ? 0x08 : 0);
	return record(RT.FONT, [
		...u16(o.height ?? 200),
		...u16(flags),
		...u16(o.color ?? 0x7fff),
		...u16(o.bold ? 700 : 400),
		...u16(o.script ?? 0),
		o.underline ?? 0,
		o.family ?? 2,
		o.charset ?? 0,
		0,
		...shortString(o.name ?? 'Arial'),
	]);
}

export const format = (id: number, code: string): number[] =>
	record(RT.FORMAT, [...u16(id), ...xlString(code)]);

export interface XfOptions {
	font?: number;
	numFmt?: number;
	style?: boolean;
	parent?: number;
	locked?: boolean;
	hidden?: boolean;
	quotePrefix?: boolean;
	/** Horizontal (0-7), wrap, vertical (0-4), rotation, indent, shrink. */
	h?: number;
	wrap?: boolean;
	v?: number;
	rotation?: number;
	indent?: number;
	shrink?: boolean;
	/** Border line styles and colors: [left, right, top, bottom] as [dg, icv]. */
	borders?: Array<[number, number]>;
	diagonal?: { dg: number; icv: number; down?: boolean; up?: boolean };
	/** Fill pattern and colors. */
	fill?: { fls: number; fore?: number; back?: number };
	/** "Used attribute" bits (byte 9). */
	used?: number;
}

export function xf(o: XfOptions = {}): number[] {
	const flags =
		(o.locked === false ? 0 : 0x01) |
		(o.hidden ? 0x02 : 0) |
		(o.style ? 0x04 : 0) |
		(o.quotePrefix ? 0x08 : 0) |
		((o.style ? 0xfff : (o.parent ?? 0)) << 4);
	const align = (o.h ?? 0) | (o.wrap ? 0x08 : 0) | ((o.v ?? 2) << 4);
	const b = o.borders ?? [
		[0, 0],
		[0, 0],
		[0, 0],
		[0, 0],
	];
	const d = o.diagonal;
	const border1 =
		(b[0][0] |
			(b[1][0] << 4) |
			(b[2][0] << 8) |
			(b[3][0] << 12) |
			(b[0][1] << 16) |
			(b[1][1] << 23)) +
		(d ? (d.down ? 1 : 0) | (d.up ? 2 : 0) : 0) * 2 ** 30;
	const fill = o.fill ?? { fls: 0 };
	const border2 =
		(b[2][1] | (b[3][1] << 7) | ((d?.icv ?? 0) << 14) | ((d?.dg ?? 0) << 21)) + fill.fls * 2 ** 26;
	return record(RT.XF, [
		...u16(o.font ?? 0),
		...u16(o.numFmt ?? 0),
		...u16(flags),
		align,
		o.rotation ?? 0,
		(o.indent ?? 0) | (o.shrink ? 0x10 : 0),
		o.used ?? 0,
		...u32(border1),
		...u32(border2),
		...u16((fill.fore ?? 64) | ((fill.back ?? 65) << 7)),
	]);
}

export const styleBuiltin = (xfIndex: number, id: number, level = 0xff): number[] =>
	record(RT.STYLE, [...u16(xfIndex | 0x8000), id, level]);

export const styleUser = (xfIndex: number, name: string): number[] =>
	record(RT.STYLE, [...u16(xfIndex), ...xlString(name)]);

export function palette(colors: string[]): number[] {
	const data = [...u16(colors.length)];
	for (const c of colors) {
		data.push(
			parseInt(c.slice(0, 2), 16),
			parseInt(c.slice(2, 4), 16),
			parseInt(c.slice(4, 6), 16),
			0,
		);
	}
	return record(RT.PALETTE, data);
}

/** Four fonts, the 15 style XFs, the default cell XF (15) and the Normal style. */
export function standardStyles(): number[][] {
	return [
		font(),
		font({ bold: true }),
		font({ italic: true }),
		font({ name: 'Courier New' }),
		...Array.from({ length: 15 }, () => xf({ style: true })),
		xf(),
		styleBuiltin(0, 0),
	];
}

export interface NameOptions {
	/** Name text, or the code of a built-in name. */
	name: string | number;
	rgce: number[];
	extra?: number[];
	/** 1-based sheet index for sheet-scoped names. */
	itab?: number;
	hidden?: boolean;
	func?: boolean;
}

export function definedName(o: NameOptions): number[] {
	const builtin = typeof o.name === 'number';
	const text = typeof o.name === 'number' ? String.fromCharCode(o.name) : o.name;
	const flags = (o.hidden ? 0x01 : 0) | (o.func ? 0x02 : 0) | (builtin ? 0x20 : 0);
	return record(RT.NAME, [
		...u16(flags),
		0,
		text.length,
		...u16(o.rgce.length),
		...u16(0),
		...u16(o.itab ?? 0),
		0,
		0,
		0,
		0,
		...stringNoCch(text),
		...o.rgce,
		...(o.extra ?? []),
	]);
}

export const supbookSelf = (sheets: number): number[] =>
	record(RT.SUPBOOK, [...u16(sheets), 0x01, 0x04]);
export const supbookAddin = (): number[] => record(RT.SUPBOOK, [...u16(1), 0x01, 0x3a]);
export const supbookExternal = (path: string, sheets: string[]): number[] =>
	record(RT.SUPBOOK, [
		...u16(sheets.length),
		...u16(path.length),
		...stringNoCch(path),
		...sheets.flatMap((s) => xlString(s)),
	]);
export const externName = (name: string): number[] =>
	record(RT.EXTERNNAME, [...u16(0), ...u32(0), ...shortString(name), ...u16(2), 0x1c, 0x17]);
export const externSheet = (xti: Array<[number, number, number]>): number[] =>
	record(RT.EXTERNSHEET, [
		...u16(xti.length),
		...xti.flatMap(([s, a, b]) => [...u16(s), ...u16(a), ...u16(b)]),
	]);

/** A single-record SST of plain strings. */
export function sst(strings: string[]): number[] {
	return record(RT.SST, [
		...u32(strings.length),
		...u32(strings.length),
		...strings.flatMap((s) => xlString(s)),
	]);
}

export const filepass = (): number[] =>
	record(RT.FILEPASS, [...u16(1), ...u16(1), ...u16(1), ...new Array<number>(48).fill(0)]);

export interface SheetSpec {
	name: string;
	records: number[][];
	state?: 'visible' | 'hidden' | 'veryHidden';
	/** BOUNDSHEET `dt` (0 worksheet, 1 macro sheet, 2 chart). */
	type?: number;
	/** BOF type of the substream. */
	bofType?: number;
}

const STATE_CODE = { visible: 0, hidden: 1, veryHidden: 2 } as const;

function boundSheet(offset: number, s: SheetSpec): number[] {
	return record(RT.BOUNDSHEET8, [
		...u32(offset),
		STATE_CODE[s.state ?? 'visible'],
		s.type ?? 0,
		...shortString(s.name),
	]);
}

/** Globals (BOF, `globals`, BOUNDSHEET8 per sheet, EOF) followed by the sheet substreams. */
export function workbookStream(
	globals: number[][],
	sheets: SheetSpec[],
	version = 0x0600,
): Uint8Array {
	const head = [bof(BOF_TYPE.globals, version), ...globals];
	const bodies = sheets.map((s) =>
		[bof(s.bofType ?? BOF_TYPE.worksheet, version), ...s.records, eof()].flat(),
	);
	let offset =
		head.flat().length + sheets.reduce((n, s) => n + boundSheet(0, s).length, 0) + eof().length;
	const directory = sheets.map((s, i) => {
		const r = boundSheet(offset, s);
		offset += bodies[i].length;
		return r;
	});
	return concat([...head, ...directory, eof(), ...bodies]);
}

/** A complete `.xls` file: the workbook stream in a compound file, plus extra streams. */
export function xlsFile(
	stream: Uint8Array,
	extra: CfbStream[] = [],
	streamName = 'Workbook',
): Uint8Array {
	return writeCompoundFileForTest([{ name: streamName, data: stream }, ...extra]);
}

/** Shorthand: standard styles, the given globals and sheets, as an `.xls` file. */
export function buildXls(globals: number[][], sheets: SheetSpec[]): Uint8Array {
	return xlsFile(workbookStream([...standardStyles(), ...globals], sheets));
}

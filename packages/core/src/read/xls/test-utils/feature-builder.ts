/** Test-only builders for hyperlink, note, validation and conditional format records. */
import { RT } from '../biff-opcodes';
import { chars, record, ref8, u16, u32, xlString } from './bytes';

const STD_LINK = [
	0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const URL_MONIKER = [
	0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const FILE_MONIKER = [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46];
const hlString = (s: string): number[] => [...u32(s.length + 1), ...chars(`${s}\u0000`, true)];

export interface LinkOptions {
	url?: string;
	file?: { path: string; up?: number };
	location?: string;
	display?: string;
}

export function hlink(ref: [number, number, number, number], o: LinkOptions): number[] {
	const flags = (o.url || o.file ? 0x03 : 0) | (o.location ? 0x08 : 0) | (o.display ? 0x14 : 0);
	const data = [...ref8(...ref), ...STD_LINK, ...u32(2), ...u32(flags)];
	if (o.display) {
		data.push(...hlString(o.display));
	}
	if (o.url) {
		const url = chars(`${o.url}\u0000`, true);
		data.push(...URL_MONIKER, ...u32(url.length), ...url);
	} else if (o.file) {
		const ansi = [...chars(o.file.path, false), 0];
		const wide = chars(o.file.path, true);
		data.push(...FILE_MONIKER, ...u16(o.file.up ?? 0), ...u32(ansi.length), ...ansi);
		data.push(...u16(0xffff), ...u16(0xdead), ...new Array<number>(20).fill(0));
		data.push(...u32(wide.length + 6), ...u32(wide.length), ...u16(3), ...wide);
	}
	if (o.location) {
		data.push(...hlString(o.location));
	}
	return record(RT.HLINK, data);
}

export const hlinkTooltip = (ref: [number, number, number, number], text: string): number[] =>
	record(RT.HLINKTOOLTIP, [
		...u16(RT.HLINKTOOLTIP),
		...ref8(...ref),
		...chars(`${text}\u0000`, true),
	]);

export const note = (
	r: number,
	c: number,
	objId: number,
	author: string,
	visible = false,
): number[] =>
	record(RT.NOTE, [
		...u16(r),
		...u16(c),
		...u16(visible ? 2 : 0),
		...u16(objId),
		...xlString(author),
		0,
	]);

/** OBJ with only the common object data (ftCmo) and the end marker. */
export const obj = (type: number, id: number): number[] =>
	record(RT.OBJ, [
		...u16(0x15),
		...u16(0x12),
		...u16(type),
		...u16(id),
		...u16(0x4011),
		...new Array<number>(12).fill(0),
		...u32(0),
	]);

/** TXO with its text CONTINUE and formatting-run CONTINUE. */
export function txo(
	text: string,
	runs: Array<[number, number]> = [[0, 0]],
	high = false,
): number[] {
	const body = [
		...u16(0x0212),
		...u16(0),
		...new Array<number>(6).fill(0),
		...u16(text.length),
		...u16((runs.length + 1) * 8),
		...u16(0),
		...u16(0),
		...u32(0),
	];
	const textData = [high ? 1 : 0, ...chars(text, high)];
	const runData = [...runs, [text.length, 0] as [number, number]].flatMap(([i, f]) => [
		...u16(i),
		...u16(f),
		...u32(0),
	]);
	return [
		...record(RT.TXO, body),
		...record(RT.CONTINUE, textData),
		...record(RT.CONTINUE, runData),
	];
}

export const dval = (count: number): number[] =>
	record(RT.DVAL, [...u16(0), ...u32(0), ...u32(0), ...u32(0xffffffff), ...u32(count)]);

export interface DvOptions {
	type: number;
	op?: number;
	errStyle?: number;
	list?: boolean;
	allowBlank?: boolean;
	suppressDropdown?: boolean;
	showInput?: boolean;
	showError?: boolean;
	promptTitle?: string;
	prompt?: string;
	errorTitle?: string;
	error?: string;
	f1?: number[];
	f2?: number[];
	ranges: Array<[number, number, number, number]>;
}

export function dv(o: DvOptions): number[] {
	const flags =
		o.type |
		((o.errStyle ?? 0) << 4) |
		(o.list ? 0x80 : 0) |
		(o.allowBlank ? 0x100 : 0) |
		(o.suppressDropdown ? 0x200 : 0) |
		(o.showInput ? 0x40000 : 0) |
		(o.showError ? 0x80000 : 0) |
		((o.op ?? 0) << 20);
	const str = (s?: string): number[] => xlString(s ?? '\u0000');
	const f = (bytes?: number[]): number[] => [
		...u16(bytes?.length ?? 0),
		...u16(0),
		...(bytes ?? []),
	];
	return record(RT.DV, [
		...u32(flags),
		...str(o.promptTitle),
		...str(o.errorTitle),
		...str(o.prompt),
		...str(o.error),
		...f(o.f1),
		...f(o.f2),
		...u16(o.ranges.length),
		...o.ranges.flatMap(([r1, r2, c1, c2]) => ref8(r1, r2, c1, c2)),
	]);
}

export function condFmt(count: number, ranges: Array<[number, number, number, number]>): number[] {
	const [r1, r2, c1, c2] = ranges[0];
	return record(RT.CONDFMT, [
		...u16(count),
		...u16(1 << 1),
		...ref8(r1, r2, c1, c2),
		...u16(ranges.length),
		...ranges.flatMap(([a, b, c, d]) => ref8(a, b, c, d)),
	]);
}

/** CF rule with an optional font color / bold / italic block and solid fill block. */
export function cf(o: {
	kind: number;
	op?: number;
	f1: number[];
	f2?: number[];
	fontColor?: number;
	bold?: boolean;
	italic?: boolean;
	fill?: number;
	border?: number;
}): number[] {
	let flags = 0x003fffff;
	const blocks: number[] = [];
	if (o.fontColor !== undefined || o.bold !== undefined || o.italic !== undefined) {
		flags |= 1 << 26;
		const fontBlock = new Array<number>(118).fill(0);
		const view = new DataView(Uint8Array.from(fontBlock).buffer);
		view.setInt32(64, -1, true);
		view.setUint32(68, o.italic ? 0x02 : 0, true);
		view.setInt16(72, o.bold ? 700 : 400, true);
		view.setInt32(80, o.fontColor ?? -1, true);
		view.setUint32(88, (o.italic === undefined ? 0x02 : 0) | 0x80, true);
		view.setUint32(92, 1, true);
		view.setUint32(96, 1, true);
		view.setUint32(100, o.bold === undefined ? 1 : 0, true);
		blocks.push(...new Uint8Array(view.buffer));
	}
	if (o.border !== undefined) {
		flags |= 1 << 28;
		flags &= ~((1 << 10) | (1 << 11) | (1 << 12) | (1 << 13));
		blocks.push(
			...u32(0x1111 | (o.border << 16) | (o.border << 23)),
			...u32(o.border | (o.border << 7)),
		);
	}
	if (o.fill !== undefined) {
		flags |= 1 << 29;
		flags &= ~((1 << 16) | (1 << 18));
		blocks.push(...u16(1 << 10), ...u16(64 | (o.fill << 7)));
	}
	const f2 = o.f2 ?? [];
	return record(RT.CF, [
		o.kind,
		o.op ?? 0,
		...u16(o.f1.length),
		...u16(f2.length),
		...u32(flags >>> 0),
		...u16(0),
		...blocks,
		...o.f1,
		...f2,
	]);
}

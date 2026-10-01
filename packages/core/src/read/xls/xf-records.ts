/**
 * Decoders for the BIFF8 formatting records: FONT, FORMAT, XF, STYLE and
 * PALETTE. Colors stay palette indices (`{ indexed }`), as in the file;
 * the PALETTE record becomes `styles.indexedColors`.
 */
import type {
	AlignmentSpec,
	BorderEdge,
	BorderSpec,
	BorderStyle,
	ColorSpec,
	FillSpec,
	FontSpec,
	HorizontalAlignment,
	PatternFill,
	PatternType,
	UnderlineStyle,
	VerticalAlignment,
} from '../../types/style';
import { BiffCursor } from './biff-cursor';
import type { BiffRecord } from './biff-records';

/** Border line styles by BIFF `dg` value. */
export const BORDER_STYLES: readonly BorderStyle[] = [
	'none',
	'thin',
	'medium',
	'dashed',
	'dotted',
	'thick',
	'double',
	'hair',
	'mediumDashed',
	'dashDot',
	'mediumDashDot',
	'dashDotDot',
	'mediumDashDotDot',
	'slantDashDot',
];

/** Fill patterns by BIFF `fls` value. */
export const PATTERN_TYPES: readonly PatternType[] = [
	'none',
	'solid',
	'mediumGray',
	'darkGray',
	'lightGray',
	'darkHorizontal',
	'darkVertical',
	'darkDown',
	'darkUp',
	'darkGrid',
	'darkTrellis',
	'lightHorizontal',
	'lightVertical',
	'lightDown',
	'lightUp',
	'lightGrid',
	'lightTrellis',
	'gray125',
	'gray0625',
];

const HORIZONTAL: readonly HorizontalAlignment[] = [
	'general',
	'left',
	'center',
	'right',
	'fill',
	'justify',
	'centerContinuous',
	'distributed',
];

const VERTICAL: readonly VerticalAlignment[] = [
	'top',
	'center',
	'bottom',
	'justify',
	'distributed',
];

/** Palette color (`icv`); 0x7FFF and other out-of-palette values mean automatic. */
export function icvColor(icv: number): ColorSpec | undefined {
	return icv >= 0 && icv <= 65 ? { indexed: icv } : undefined;
}

function underline(uls: number): UnderlineStyle | undefined {
	switch (uls) {
		case 0x01:
			return 'single';
		case 0x02:
			return 'double';
		case 0x21:
			return 'singleAccounting';
		case 0x22:
			return 'doubleAccounting';
		default:
			return undefined;
	}
}

/** FONT record. */
export function parseFont(record: BiffRecord): FontSpec {
	const cur = new BiffCursor(record.data, record.breaks);
	const height = cur.u16();
	const flags = cur.u16();
	const icv = cur.u16();
	const weight = cur.u16();
	const script = cur.u16();
	const uls = cur.u8();
	const family = cur.u8();
	const charset = cur.u8();
	cur.skip(1);
	const font: FontSpec = { name: cur.shortString(), size: height / 20 };
	if (weight >= 600) {
		font.bold = true;
	}
	if (flags & 0x02) {
		font.italic = true;
	}
	const u = underline(uls);
	if (u) {
		font.underline = u;
	}
	if (flags & 0x08) {
		font.strike = true;
	}
	if (flags & 0x10) {
		font.outline = true;
	}
	if (flags & 0x20) {
		font.shadow = true;
	}
	if (flags & 0x40) {
		font.condense = true;
	}
	if (flags & 0x80) {
		font.extend = true;
	}
	const color = icvColor(icv);
	if (color) {
		font.color = color;
	}
	if (script === 1 || script === 2) {
		font.vertAlign = script === 1 ? 'superscript' : 'subscript';
	}
	if (family > 0 && family <= 5) {
		font.family = family;
	}
	if (charset > 0) {
		font.charset = charset;
	}
	return font;
}

/** FORMAT record: number format id and code. */
export function parseFormat(record: BiffRecord): { id: number; code: string } {
	const cur = new BiffCursor(record.data, record.breaks);
	const id = cur.u16();
	return { id, code: cur.unicodeString() };
}

/** A decoded XF record (cell or style format). */
export interface RawXf {
	/** BIFF font index (index 4 does not exist). */
	font: number;
	numFmtId: number;
	isStyle: boolean;
	/** Parent style XF (cell XFs). */
	parent: number;
	locked: boolean;
	hidden: boolean;
	quotePrefix: boolean;
	alignment: AlignmentSpec;
	border: BorderSpec;
	fill: FillSpec;
	/** "Attribute group used" bits: number format, font, alignment, border, fill, protection. */
	used: {
		num: boolean;
		font: boolean;
		align: boolean;
		border: boolean;
		fill: boolean;
		prot: boolean;
	};
}

function edge(dg: number, icv: number): BorderEdge | undefined {
	const style = BORDER_STYLES[dg];
	if (!style || style === 'none') {
		return undefined;
	}
	const color = icvColor(icv);
	return color ? { style, color } : { style };
}

/** Pattern fill from `fls` and its two palette colors. */
function patternFill(fls: number, fore: number, back: number): FillSpec {
	const patternType = PATTERN_TYPES[fls] ?? 'none';
	if (patternType === 'none') {
		return { type: 'pattern', patternType };
	}
	const fill: PatternFill = { type: 'pattern', patternType };
	const fg = icvColor(fore);
	const bg = icvColor(back);
	if (fg) {
		fill.fgColor = fg;
	}
	if (bg) {
		fill.bgColor = bg;
	}
	return fill;
}

function alignment(b6: number, trot: number, b8: number): AlignmentSpec {
	const a: AlignmentSpec = {};
	const h = HORIZONTAL[b6 & 0x07];
	if (h && h !== 'general') {
		a.horizontal = h;
	}
	if (b6 & 0x08) {
		a.wrapText = true;
	}
	const v = VERTICAL[(b6 >> 4) & 0x07];
	if (v && v !== 'bottom') {
		a.vertical = v;
	}
	if (b6 & 0x80) {
		a.justifyLastLine = true;
	}
	if (trot > 0 && (trot <= 180 || trot === 255)) {
		a.textRotation = trot;
	}
	if (b8 & 0x0f) {
		a.indent = b8 & 0x0f;
	}
	if (b8 & 0x10) {
		a.shrinkToFit = true;
	}
	const order = (b8 >> 6) & 0x03;
	if (order === 1 || order === 2) {
		a.readingOrder = order;
	}
	return a;
}

/** XF record (BIFF8, 20 bytes). */
export function parseXf(record: BiffRecord): RawXf {
	const cur = new BiffCursor(record.data, record.breaks);
	const font = cur.u16();
	const numFmtId = cur.u16();
	const flags = cur.u16();
	const b6 = cur.u8();
	const trot = cur.u8();
	const b8 = cur.u8();
	const used = cur.u8();
	const border1 = cur.u32();
	const border2 = cur.u32();
	const colors = cur.u16();
	const border: BorderSpec = {};
	const sides: Array<['left' | 'right' | 'top' | 'bottom', number, number]> = [
		['left', border1 & 0x0f, (border1 >>> 16) & 0x7f],
		['right', (border1 >>> 4) & 0x0f, (border1 >>> 23) & 0x7f],
		['top', (border1 >>> 8) & 0x0f, border2 & 0x7f],
		['bottom', (border1 >>> 12) & 0x0f, (border2 >>> 7) & 0x7f],
	];
	for (const [side, dg, icv] of sides) {
		const e = edge(dg, icv);
		if (e) {
			border[side] = e;
		}
	}
	const diag = (border1 >>> 30) & 0x03;
	const diagonal = diag ? edge((border2 >>> 21) & 0x0f, (border2 >>> 14) & 0x7f) : undefined;
	if (diagonal) {
		border.diagonal = diagonal;
		if (diag & 1) {
			border.diagonalDown = true;
		}
		if (diag & 2) {
			border.diagonalUp = true;
		}
	}
	return {
		font,
		numFmtId,
		isStyle: (flags & 0x04) !== 0,
		parent: flags >>> 4,
		locked: (flags & 0x01) !== 0,
		hidden: (flags & 0x02) !== 0,
		quotePrefix: (flags & 0x08) !== 0,
		alignment: alignment(b6, trot, b8),
		border,
		fill: patternFill((border2 >>> 26) & 0x3f, colors & 0x7f, (colors >>> 7) & 0x7f),
		used: {
			num: (used & 0x04) !== 0,
			font: (used & 0x08) !== 0,
			align: (used & 0x10) !== 0,
			border: (used & 0x20) !== 0,
			fill: (used & 0x40) !== 0,
			prot: (used & 0x80) !== 0,
		},
	};
}

/** A STYLE record: a built-in style (by id) or a user-defined style name. */
export interface RawStyle {
	xf: number;
	builtinId?: number;
	level?: number;
	name?: string;
}

export function parseStyle(record: BiffRecord): RawStyle {
	const cur = new BiffCursor(record.data, record.breaks);
	const v = cur.u16();
	const xf = v & 0x0fff;
	if (v & 0x8000) {
		return { xf, builtinId: cur.u8(), level: cur.u8() };
	}
	return { xf, name: cur.unicodeString() };
}

/** PALETTE record: `RRGGBB` colors for palette indices 8 and up. */
export function parsePalette(record: BiffRecord): string[] {
	const cur = new BiffCursor(record.data, record.breaks);
	const count = Math.min(cur.u16(), Math.floor(cur.remaining / 4), 56);
	const hex = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
	const out: string[] = [];
	for (let i = 0; i < count; i++) {
		const [r, g, b] = [cur.u8(), cur.u8(), cur.u8()];
		cur.skip(1);
		out.push(hex(r) + hex(g) + hex(b));
	}
	return out;
}

/**
 * DXFN ([MS-XLS]): the differential format of a CF record. A
 * 6-byte header of "not specified" (Ninch) flags and block-present flags,
 * then optional number format, font, alignment, border, pattern and
 * protection blocks. Only the attributes the rule specifies are set.
 */
import { numberFormatCode } from '../../styles/builtin-number-formats';
import type {
	AlignmentSpec,
	BorderEdge,
	BorderSpec,
	DifferentialFormat,
	FontSpec,
	PatternFill,
} from '../../types/style';
import type { BiffCursor } from './biff-cursor';
import { BORDER_STYLES, icvColor, PATTERN_TYPES } from './xf-records';

const HORIZONTAL = [
	'general',
	'left',
	'center',
	'right',
	'fill',
	'justify',
	'centerContinuous',
	'distributed',
] as const;
const VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'] as const;

const bit = (flags: number, n: number): boolean => ((flags >>> n) & 1) === 1;

function readFont(cur: BiffCursor): FontSpec {
	const start = cur.pos;
	const cchFont = cur.u8();
	const name = cchFont > 0 && cchFont <= 63 ? cur.stringNoCch(cchFont) : '';
	cur.pos = start + 64;
	const height = cur.i32();
	const ts = cur.u32();
	const weight = cur.i16();
	const script = cur.i16();
	const uls = cur.u8();
	cur.skip(3);
	const color = cur.i32();
	cur.skip(4);
	const tsNinch = cur.u32();
	const sssNinch = cur.u32();
	const ulsNinch = cur.u32();
	const blsNinch = cur.u32();
	cur.skip(4 + 4 + 4 + 2);
	const font: FontSpec = {};
	if (name) {
		font.name = name;
	}
	if (height > 0) {
		font.size = height / 20;
	}
	if (!bit(tsNinch, 1)) {
		font.italic = bit(ts, 1);
	}
	if (!bit(tsNinch, 7)) {
		font.strike = bit(ts, 7);
	}
	// Excel sets bold and italic together (one "font style" choice), flagged by the italic bit.
	if ((blsNinch === 0 || !bit(tsNinch, 1)) && weight > 0 && weight < 0x7fff) {
		font.bold = weight >= 600;
	}
	if (sssNinch === 0 && (script === 1 || script === 2)) {
		font.vertAlign = script === 1 ? 'superscript' : 'subscript';
	}
	if (ulsNinch === 0) {
		const map: Record<number, FontSpec['underline']> = {
			0x00: 'none',
			0x01: 'single',
			0x02: 'double',
			0x21: 'singleAccounting',
			0x22: 'doubleAccounting',
		};
		const u = map[uls];
		if (u) {
			font.underline = u;
		}
	}
	const c = color >= 0 ? icvColor(color) : undefined;
	if (c) {
		font.color = c;
	}
	return font;
}

function readAlignment(cur: BiffCursor, flags: number): AlignmentSpec {
	const v = cur.u32();
	const indent = cur.i32();
	const a: AlignmentSpec = {};
	if (!bit(flags, 0)) {
		a.horizontal = HORIZONTAL[v & 0x07];
	}
	if (!bit(flags, 2)) {
		a.wrapText = bit(v, 3);
	}
	if (!bit(flags, 1)) {
		a.vertical = VERTICAL[(v >>> 4) & 0x07] ?? 'bottom';
	}
	if (!bit(flags, 3)) {
		a.textRotation = (v >>> 8) & 0xff;
	}
	if (!bit(flags, 5)) {
		a.indent = indent >= 0 && indent <= 250 ? indent : (v >>> 16) & 0x0f;
	}
	if (!bit(flags, 6)) {
		a.shrinkToFit = bit(v, 20);
	}
	return a;
}

function readBorder(cur: BiffCursor, flags: number): BorderSpec {
	const a = cur.u32();
	const b = cur.u32();
	const edge = (dg: number, icv: number): BorderEdge => {
		const style = BORDER_STYLES[dg] ?? 'none';
		const color = style === 'none' ? undefined : icvColor(icv);
		return color ? { style, color } : { style };
	};
	const border: BorderSpec = {};
	if (!bit(flags, 10)) {
		border.left = edge(a & 0x0f, (a >>> 16) & 0x7f);
	}
	if (!bit(flags, 11)) {
		border.right = edge((a >>> 4) & 0x0f, (a >>> 23) & 0x7f);
	}
	if (!bit(flags, 12)) {
		border.top = edge((a >>> 8) & 0x0f, b & 0x7f);
	}
	if (!bit(flags, 13)) {
		border.bottom = edge((a >>> 12) & 0x0f, (b >>> 7) & 0x7f);
	}
	const down = !bit(flags, 14) && bit(a, 30);
	const up = !bit(flags, 15) && bit(a, 31);
	if (down || up) {
		border.diagonal = edge((b >>> 21) & 0x0f, (b >>> 14) & 0x7f);
		border.diagonalDown = down || undefined;
		border.diagonalUp = up || undefined;
	}
	return border;
}

/** Solid CF fills keep their color in the pattern background, like SpreadsheetML `dxf` fills. */
function readPattern(cur: BiffCursor, flags: number): PatternFill | undefined {
	const fls = (cur.u16() >>> 10) & 0x3f;
	const colors = cur.u16();
	const fg = bit(flags, 17) ? undefined : icvColor(colors & 0x7f);
	const bg = bit(flags, 18) ? undefined : icvColor((colors >>> 7) & 0x7f);
	const patternType = bit(flags, 16) ? (fg || bg ? 'solid' : undefined) : PATTERN_TYPES[fls];
	if (!patternType) {
		return undefined;
	}
	const fill: PatternFill = { type: 'pattern', patternType };
	const back = patternType === 'solid' ? (bg ?? fg) : bg;
	if (fg && patternType !== 'solid') {
		fill.fgColor = fg;
	}
	if (back) {
		fill.bgColor = back;
	}
	return fill;
}

/**
 * Read a DXFN at the cursor. `customFormat` returns the id to use for a
 * custom number format code; built-in ids resolve through `numFmts`.
 */
export function readDxfn(
	cur: BiffCursor,
	numFmts: Readonly<Record<number, string>>,
	customFormat: (code: string) => number,
): DifferentialFormat {
	const flags = cur.u32();
	const flags2 = cur.u16();
	const dxf: DifferentialFormat = {};
	if (bit(flags, 25)) {
		if (bit(flags2, 0)) {
			const start = cur.pos;
			const cb = cur.u16();
			const code = cur.unicodeString();
			cur.pos = Math.max(cur.pos, start + cb);
			dxf.numFmt = { id: customFormat(code), code };
		} else {
			cur.skip(1);
			const id = cur.u8();
			if (!bit(flags, 19)) {
				dxf.numFmt = { id, code: numberFormatCode(id, numFmts) };
			}
		}
	}
	if (bit(flags, 26)) {
		const font = readFont(cur);
		if (Object.keys(font).length) {
			dxf.font = font;
		}
	}
	if (bit(flags, 27)) {
		dxf.alignment = readAlignment(cur, flags);
	}
	if (bit(flags, 28)) {
		dxf.border = readBorder(cur, flags);
	}
	if (bit(flags, 29)) {
		const fill = readPattern(cur, flags);
		if (fill) {
			dxf.fill = fill;
		}
	}
	if (bit(flags, 30)) {
		const prot = cur.u16();
		dxf.protection = {};
		if (!bit(flags, 8)) {
			dxf.protection.locked = bit(prot, 0);
		}
		if (!bit(flags, 9)) {
			dxf.protection.hidden = bit(prot, 1);
		}
	}
	return dxf;
}

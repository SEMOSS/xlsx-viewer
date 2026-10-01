/**
 * Legacy VML drawings (`xl/drawings/vmlDrawingN.vml`) hold the note boxes of
 * legacy comments. VML written by older Excel versions is not always
 * well-formed XML, so note shapes are located with a tolerant scanner.
 */

export interface VmlNoteShape {
	/** 0-based row / column of the commented cell (`x:Row` / `x:Column`). */
	row: number;
	col: number;
	/** `<x:Visible/>`: the note is always shown. */
	visible: boolean;
	/** Box size in pixels from the shape style, when given. */
	width?: number;
	height?: number;
}

const SHAPE_RE = /<(?:[\w.-]+:)?shape\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?shape\s*>/gi;
const STYLE_RE = /\bstyle\s*=\s*(?:"([^"]*)"|'([^']*)')/i;
const CLIENT_DATA_RE = /<(?:[\w.-]+:)?ClientData\b[^>]*\bObjectType\s*=\s*["']([^"']*)["']/i;
const ROW_RE = /<(?:[\w.-]+:)?Row\s*>\s*(\d+)\s*</i;
const COL_RE = /<(?:[\w.-]+:)?Column\s*>\s*(\d+)\s*</i;
const VISIBLE_RE = /<(?:[\w.-]+:)?Visible\b\s*(\/>|>([^<]*)<)/i;

/** CSS length in a VML style -> pixels at 96 DPI. */
export function cssLengthToPx(value: string): number | undefined {
	const m = /^\s*(-?\d*\.?\d+)\s*(pt|px|in|cm|mm|pc|em)?\s*$/i.exec(value);
	if (!m) {
		return undefined;
	}
	const n = Number(m[1]);
	switch ((m[2] ?? 'px').toLowerCase()) {
		case 'pt':
			return (n * 96) / 72;
		case 'in':
			return n * 96;
		case 'cm':
			return (n * 96) / 2.54;
		case 'mm':
			return (n * 96) / 25.4;
		case 'pc':
			return n * 16;
		case 'em':
			return n * 16;
		default:
			return n;
	}
}

function styleProps(style: string): Map<string, string> {
	const out = new Map<string, string>();
	for (const part of style.split(';')) {
		const i = part.indexOf(':');
		if (i > 0) {
			out.set(part.slice(0, i).trim().toLowerCase(), part.slice(i + 1).trim());
		}
	}
	return out;
}

/** Note shapes (`ObjectType="Note"`) of a VML drawing part. */
export function parseVmlNoteShapes(vml: string): VmlNoteShape[] {
	const out: VmlNoteShape[] = [];
	SHAPE_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	while ((m = SHAPE_RE.exec(vml)) !== null) {
		if (m[1].trimEnd().endsWith('/')) {
			// Self-closing shape: rescan right after it so the next shape is not swallowed.
			SHAPE_RE.lastIndex = m.index + m[0].indexOf('>') + 1;
			continue;
		}
		const body = m[2];
		const client = CLIENT_DATA_RE.exec(body);
		if (!client || client[1].toLowerCase() !== 'note') {
			continue;
		}
		const row = ROW_RE.exec(body);
		const col = COL_RE.exec(body);
		if (!row || !col) {
			continue;
		}
		const visibleMatch = VISIBLE_RE.exec(body);
		const visibleText = visibleMatch?.[2]?.trim().toLowerCase();
		const shape: VmlNoteShape = {
			row: Number(row[1]),
			col: Number(col[1]),
			visible: !!visibleMatch && visibleText !== 'false' && visibleText !== '0',
		};
		const styleMatch = STYLE_RE.exec(m[1]);
		if (styleMatch) {
			const props = styleProps(styleMatch[1] ?? styleMatch[2] ?? '');
			const width = cssLengthToPx(props.get('width') ?? '');
			const height = cssLengthToPx(props.get('height') ?? '');
			if (width !== undefined && width > 0) {
				shape.width = Math.round(width * 100) / 100;
			}
			if (height !== undefined && height > 0) {
				shape.height = Math.round(height * 100) / 100;
			}
		}
		out.push(shape);
	}
	return out;
}

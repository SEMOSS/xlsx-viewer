/**
 * Readers for the DrawingML fragments embedded in chart parts: colors, fills
 * and outlines (text lives in `text-read.ts`). Elements are matched by local
 * name so files that use unusual namespace prefixes still parse.
 */
import type { DrawingColor, ShapeFill, ShapeLine } from '../../types/drawing';
import { childLocal, childrenLocal, isElement, localName } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';

// ── Generic `val` helpers ──────────────────────────────────────────────────

/** First child with the local name. */
export function kid(el: XmlElement | undefined, local: string): XmlElement | undefined {
	return childLocal(el, local);
}

/** All children with the local name. */
export function kids(el: XmlElement | undefined, local: string): XmlElement[] {
	return childrenLocal(el, local);
}

/** `val` attribute of a child element. */
export function valOf(el: XmlElement | undefined, local: string): string | undefined {
	return childLocal(el, local)?.attrs.val;
}

/** Numeric `val` of a child element. */
export function numOf(el: XmlElement | undefined, local: string): number | undefined {
	const v = valOf(el, local);
	if (v === undefined || v.trim() === '') {
		return undefined;
	}
	const n = Number(v);
	return Number.isFinite(n) ? n : undefined;
}

/** Parse a CT_Boolean-style value (`1`, `true`, `0`, `false`); a missing `val` means true. */
export function parseBool(v: string | undefined): boolean {
	if (v === undefined) {
		return true;
	}
	const s = v.trim().toLowerCase();
	return !(s === '0' || s === 'false' || s === 'off');
}

/** CT_Boolean child: absent is `undefined`, present without `val` is `true`. */
export function boolOf(el: XmlElement | undefined, local: string): boolean | undefined {
	const c = childLocal(el, local);
	return c ? parseBool(c.attrs.val) : undefined;
}

/** Numeric attribute. */
export function numAttr(el: XmlElement | undefined, name: string): number | undefined {
	const v = el?.attrs[name];
	if (v === undefined || v.trim() === '') {
		return undefined;
	}
	const n = Number(v);
	return Number.isFinite(n) ? n : undefined;
}

// ── Colors ─────────────────────────────────────────────────────────────────

const COLOR_KINDS: Record<string, DrawingColor['kind']> = {
	srgbClr: 'srgb',
	schemeClr: 'scheme',
	sysClr: 'sys',
	prstClr: 'preset',
	hslClr: 'hsl',
	scrgbClr: 'scrgb',
};

/** Read one color element (`a:srgbClr`, `a:schemeClr`, ...). */
export function readColorElement(el: XmlElement): DrawingColor | undefined {
	const kind = COLOR_KINDS[localName(el.name)];
	if (!kind) {
		return undefined;
	}
	let value: string;
	switch (kind) {
		case 'sys':
			value = el.attrs.lastClr ?? el.attrs.val ?? '000000';
			break;
		case 'hsl':
			value = `${el.attrs.hue ?? 0},${el.attrs.sat ?? 0},${el.attrs.lum ?? 0}`;
			break;
		case 'scrgb':
			value = `${el.attrs.r ?? 0},${el.attrs.g ?? 0},${el.attrs.b ?? 0}`;
			break;
		default:
			value = el.attrs.val ?? '';
	}
	const color: DrawingColor = { kind, value };
	const transforms: Array<{ name: string; val?: number }> = [];
	for (const t of el.children) {
		if (!isElement(t)) {
			continue;
		}
		const v = numAttr(t, 'val');
		transforms.push(
			v === undefined ? { name: localName(t.name) } : { name: localName(t.name), val: v },
		);
	}
	if (transforms.length) {
		color.transforms = transforms;
	}
	return color;
}

/** The first color child of `parent` (e.g. the content of `a:solidFill`). */
export function readColor(parent: XmlElement | undefined): DrawingColor | undefined {
	if (!parent) {
		return undefined;
	}
	for (const c of parent.children) {
		if (isElement(c) && COLOR_KINDS[localName(c.name)]) {
			return readColorElement(c);
		}
	}
	return undefined;
}

// ── Fills and lines ────────────────────────────────────────────────────────

/** Read a fill element (`a:noFill`, `a:solidFill`, `a:gradFill`, `a:pattFill`, `a:blipFill`). */
export function readFillElement(el: XmlElement): ShapeFill | undefined {
	switch (localName(el.name)) {
		case 'noFill':
			return { type: 'none' };
		case 'solidFill': {
			const color = readColor(el);
			return color ? { type: 'solid', color } : undefined;
		}
		case 'gradFill': {
			const stops: Array<{ pos: number; color: DrawingColor }> = [];
			for (const gs of kids(kid(el, 'gsLst'), 'gs')) {
				const color = readColor(gs);
				if (color) {
					stops.push({ pos: (numAttr(gs, 'pos') ?? 0) / 100000, color });
				}
			}
			const fill: ShapeFill = { type: 'gradient', stops };
			const lin = kid(el, 'lin');
			const path = kid(el, 'path');
			if (lin) {
				fill.angle = (numAttr(lin, 'ang') ?? 0) / 60000;
			} else if (path) {
				const p = path.attrs.path;
				fill.path = p === 'circle' || p === 'rect' ? p : 'shape';
			}
			return stops.length ? fill : undefined;
		}
		case 'pattFill': {
			const fg = readColor(kid(el, 'fgClr'));
			const bg = readColor(kid(el, 'bgClr'));
			const fill: ShapeFill = { type: 'pattern', preset: el.attrs.prst ?? 'pct5' };
			if (fg) {
				fill.fg = fg;
			}
			if (bg) {
				fill.bg = bg;
			}
			return fill;
		}
		case 'blipFill': {
			const blip = kid(el, 'blip');
			return { type: 'picture', mediaPath: blip?.attrs['r:embed'] ?? '' };
		}
		default:
			return undefined;
	}
}

const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill']);

/** The fill among the direct children of `parent`. */
export function readFill(parent: XmlElement | undefined): ShapeFill | undefined {
	if (!parent) {
		return undefined;
	}
	for (const c of parent.children) {
		if (isElement(c) && FILL_TAGS.has(localName(c.name))) {
			return readFillElement(c);
		}
	}
	return undefined;
}

/** Read an `a:ln` element. */
export function readLine(ln: XmlElement | undefined): ShapeLine | undefined {
	if (!ln) {
		return undefined;
	}
	const line: ShapeLine = {};
	const w = numAttr(ln, 'w');
	if (w !== undefined) {
		line.width = w;
	}
	const cap = ln.attrs.cap;
	if (cap === 'rnd' || cap === 'sq' || cap === 'flat') {
		line.cap = cap;
	}
	const fill = readFill(ln);
	if (fill) {
		line.fill = fill;
	}
	const dash = valOf(ln, 'prstDash');
	if (dash) {
		line.dash = dash;
	} else if (kid(ln, 'custDash')) {
		line.dash = 'dash';
	}
	for (const end of ['headEnd', 'tailEnd'] as const) {
		const e = kid(ln, end);
		if (e?.attrs.type && e.attrs.type !== 'none') {
			const spec: { type: string; w?: string; len?: string } = { type: e.attrs.type };
			if (e.attrs.w) {
				spec.w = e.attrs.w;
			}
			if (e.attrs.len) {
				spec.len = e.attrs.len;
			}
			line[end] = spec;
		}
	}
	return line;
}

/** Fill and outline of a `c:spPr`. */
export function readShapeProps(spPr: XmlElement | undefined): {
	fill?: ShapeFill;
	line?: ShapeLine;
} {
	const out: { fill?: ShapeFill; line?: ShapeLine } = {};
	if (!spPr) {
		return out;
	}
	const fill = readFill(spPr);
	if (fill) {
		out.fill = fill;
	}
	const line = readLine(kid(spPr, 'ln'));
	if (line) {
		out.line = line;
	}
	return out;
}

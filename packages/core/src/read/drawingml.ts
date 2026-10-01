import type { DrawingColor, ShapeFill, ShapeLine } from '../types/drawing';
import { isElement } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { kid, kids, lname, oneOf, relAttr, str } from './xml-utils';

/**
 * DrawingML primitives shared by drawing objects: colors, fills, outlines
 * and transforms (text bodies live in `drawing-text.ts`). Elements are
 * matched by local name, so any prefix (and the Strict namespaces) work.
 */

const COLOR_KINDS: Record<string, DrawingColor['kind']> = {
	srgbClr: 'srgb',
	schemeClr: 'scheme',
	sysClr: 'sys',
	prstClr: 'preset',
	hslClr: 'hsl',
	scrgbClr: 'scrgb',
};

/** Parse an ST_Percentage / ST_Angle style value: plain integers or Strict `NN%` strings. */
export function drawingNumber(value: string | undefined): number | undefined {
	if (value === undefined || value === '') {
		return undefined;
	}
	const v = value.trim();
	if (v.endsWith('%')) {
		const n = Number(v.slice(0, -1));
		return Number.isFinite(n) ? Math.round(n * 1000) : undefined;
	}
	const n = Number(v);
	return Number.isFinite(n) ? n : undefined;
}

const EMU_PER_UNIT: Record<string, number> = {
	in: 914400,
	cm: 360000,
	mm: 36000,
	pt: 12700,
	pc: 152400,
	pi: 152400,
};

/** ST_Coordinate: EMU integers, or (Strict / universal measure) values such as `2.5cm`. */
export function coordinate(value: string | undefined): number | undefined {
	if (value === undefined) {
		return undefined;
	}
	const v = value.trim();
	const m = /^(-?\d*\.?\d+)\s*(in|cm|mm|pt|pc|pi)$/i.exec(v);
	if (m) {
		return Math.round(Number(m[1]) * EMU_PER_UNIT[m[2].toLowerCase()]);
	}
	const n = Number(v);
	return v !== '' && Number.isFinite(n) ? n : undefined;
}

/** Numeric components joined with commas; percentages are normalised to 1000ths of a percent. */
function components(el: XmlElement, names: readonly string[]): string {
	return names.map((n) => drawingNumber(el.attrs[n]) ?? 0).join(',');
}

/**
 * A color element (`a:srgbClr`, `a:schemeClr`, ...) with its transforms.
 * `hslClr` values are `hue,sat,lum` (60000ths of a degree, 1000ths of a
 * percent) and `scrgbClr` values `r,g,b` (1000ths of a percent).
 */
export function parseDrawingColor(el: XmlElement | undefined): DrawingColor | undefined {
	if (!el) {
		return undefined;
	}
	const kind = COLOR_KINDS[lname(el)];
	if (!kind) {
		return undefined;
	}
	let value: string;
	switch (kind) {
		case 'sys':
			value = str(el, 'lastClr') ?? str(el, 'val') ?? '000000';
			break;
		case 'hsl':
			value = components(el, ['hue', 'sat', 'lum']);
			break;
		case 'scrgb':
			value = components(el, ['r', 'g', 'b']);
			break;
		default:
			value = str(el, 'val') ?? '';
	}
	const color: DrawingColor = { kind, value };
	const transforms: NonNullable<DrawingColor['transforms']> = [];
	for (const c of el.children) {
		if (isElement(c)) {
			const val = drawingNumber(c.attrs.val);
			transforms.push(val === undefined ? { name: lname(c) } : { name: lname(c), val });
		}
	}
	if (transforms.length > 0) {
		color.transforms = transforms;
	}
	return color;
}

/** First color child of a container (`a:solidFill`, `a:fgClr`, `a:fillRef`, ...). */
export function findColor(parent: XmlElement | undefined): DrawingColor | undefined {
	if (!parent) {
		return undefined;
	}
	for (const c of parent.children) {
		if (isElement(c) && COLOR_KINDS[lname(c)]) {
			return parseDrawingColor(c);
		}
	}
	return undefined;
}

const FILL_NAMES: ReadonlySet<string> = new Set([
	'noFill',
	'solidFill',
	'gradFill',
	'pattFill',
	'blipFill',
	'grpFill',
]);

/** Resolves a blip `r:embed` to a media path (and loads the image). */
export type BlipResolver = (relId: string) => string | undefined;

/** Parse one fill element. */
export function parseFillElement(
	el: XmlElement,
	resolveBlip?: BlipResolver,
): ShapeFill | undefined {
	switch (lname(el)) {
		case 'noFill':
			return { type: 'none' };
		case 'solidFill': {
			const color = findColor(el);
			return color ? { type: 'solid', color } : undefined;
		}
		case 'gradFill': {
			const stops: Array<{ pos: number; color: DrawingColor }> = [];
			for (const gs of kids(kid(el, 'gsLst'), 'gs')) {
				const color = findColor(gs);
				if (color) {
					stops.push({ pos: (drawingNumber(gs.attrs.pos) ?? 0) / 100000, color });
				}
			}
			const fill: Extract<ShapeFill, { type: 'gradient' }> = { type: 'gradient', stops };
			const lin = kid(el, 'lin');
			const ang = drawingNumber(lin?.attrs.ang);
			if (ang !== undefined) {
				fill.angle = ang / 60000;
			}
			const path = oneOf(str(kid(el, 'path'), 'path'), ['circle', 'rect', 'shape'] as const);
			if (path) {
				fill.path = path;
			}
			return fill;
		}
		case 'pattFill': {
			const fill: Extract<ShapeFill, { type: 'pattern' }> = {
				type: 'pattern',
				preset: str(el, 'prst') ?? 'pct5',
			};
			const fg = findColor(kid(el, 'fgClr'));
			if (fg) {
				fill.fg = fg;
			}
			const bg = findColor(kid(el, 'bgClr'));
			if (bg) {
				fill.bg = bg;
			}
			return fill;
		}
		case 'blipFill': {
			const embed = relAttr(kid(el, 'blip'), 'embed');
			const mediaPath = embed && resolveBlip ? resolveBlip(embed) : undefined;
			return mediaPath ? { type: 'picture', mediaPath } : undefined;
		}
		default:
			return undefined;
	}
}

/** The fill among the children of a properties element (`spPr`, `ln`, ...), if any. */
export function findFill(
	parent: XmlElement | undefined,
	resolveBlip?: BlipResolver,
): ShapeFill | undefined {
	if (!parent) {
		return undefined;
	}
	for (const c of parent.children) {
		if (isElement(c) && FILL_NAMES.has(lname(c))) {
			return parseFillElement(c, resolveBlip);
		}
	}
	return undefined;
}

/** True when the properties element names a fill explicitly (even an unsupported one). */
export function hasFillElement(parent: XmlElement | undefined): boolean {
	return !!parent && parent.children.some((c) => isElement(c) && FILL_NAMES.has(lname(c)));
}

function lineEnd(el: XmlElement | undefined): ShapeLine['headEnd'] {
	const type = str(el, 'type');
	if (!el || !type) {
		return undefined;
	}
	const end: NonNullable<ShapeLine['headEnd']> = { type };
	const w = str(el, 'w');
	if (w) {
		end.w = w;
	}
	const len = str(el, 'len');
	if (len) {
		end.len = len;
	}
	return end;
}

/** `a:ln` outline. */
export function parseLine(el: XmlElement | undefined): ShapeLine | undefined {
	if (!el) {
		return undefined;
	}
	const line: ShapeLine = {};
	const fill = findFill(el);
	if (fill) {
		line.fill = fill;
	}
	const w = coordinate(str(el, 'w'));
	if (w !== undefined) {
		line.width = w;
	}
	const dash = str(kid(el, 'prstDash'), 'val');
	if (dash) {
		line.dash = dash;
	}
	const cap = oneOf(str(el, 'cap'), ['rnd', 'sq', 'flat'] as const);
	if (cap) {
		line.cap = cap;
	}
	const head = lineEnd(kid(el, 'headEnd'));
	if (head) {
		line.headEnd = head;
	}
	const tail = lineEnd(kid(el, 'tailEnd'));
	if (tail) {
		line.tailEnd = tail;
	}
	return line;
}

/** Transform (`a:xfrm` / `xdr:xfrm`): rotation in degrees, flips, offset and extents in EMU. */
export interface Xfrm {
	rotation?: number;
	flipH?: boolean;
	flipV?: boolean;
	off?: { x: number; y: number };
	ext?: { cx: number; cy: number };
	chOff?: { x: number; y: number };
	chExt?: { cx: number; cy: number };
}

function point(el: XmlElement | undefined): { x: number; y: number } | undefined {
	return el ? { x: coordinate(el.attrs.x) ?? 0, y: coordinate(el.attrs.y) ?? 0 } : undefined;
}

function size(el: XmlElement | undefined): { cx: number; cy: number } | undefined {
	return el ? { cx: coordinate(el.attrs.cx) ?? 0, cy: coordinate(el.attrs.cy) ?? 0 } : undefined;
}

export function parseXfrm(el: XmlElement | undefined): Xfrm {
	if (!el) {
		return {};
	}
	const x: Xfrm = {};
	const rot = drawingNumber(el.attrs.rot);
	if (rot) {
		x.rotation = rot / 60000;
	}
	if (el.attrs.flipH === '1' || el.attrs.flipH === 'true') {
		x.flipH = true;
	}
	if (el.attrs.flipV === '1' || el.attrs.flipV === 'true') {
		x.flipV = true;
	}
	const off = point(kid(el, 'off'));
	if (off) {
		x.off = off;
	}
	const ext = size(kid(el, 'ext'));
	if (ext) {
		x.ext = ext;
	}
	const chOff = point(kid(el, 'chOff'));
	if (chOff) {
		x.chOff = chOff;
	}
	const chExt = size(kid(el, 'chExt'));
	if (chExt) {
		x.chExt = chExt;
	}
	return x;
}

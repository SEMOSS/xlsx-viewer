import type {
	DrawingAnchor,
	DrawingColor,
	DrawingObjectBase,
	PictureObject,
	ShapeFill,
	ShapeLine,
	ShapeObject,
} from '../types/drawing';
import type { XmlElement } from '../xml/types';
import type { DrawingReader } from './drawing-reader';
import { parseTextBody } from './drawing-text';
import {
	drawingNumber,
	findColor,
	findFill,
	hasFillElement,
	parseLine,
	parseXfrm,
} from './drawingml';
import { kid, kids, relAttr, str } from './xml-utils';

/** Builders for shapes, connectors and pictures (`xdr:sp`, `xdr:cxnSp`, `xdr:pic`). */

/** Default outline widths (EMU) of the Office theme's line styles, by `lnRef` index. */
const THEME_LINE_WIDTHS = [0, 6350, 12700, 19050];

/** Common object fields from `cNvPr` and the shape transform. */
export function baseFields(
	reader: DrawingReader,
	cNvPr: XmlElement | undefined,
	anchor: DrawingAnchor,
	raw: XmlElement,
	xfrm?: XmlElement,
): DrawingObjectBase {
	const base: DrawingObjectBase = {
		id: reader.uniqueId(str(cNvPr, 'id')),
		anchor,
		raw,
		dirty: false,
	};
	const name = str(cNvPr, 'name');
	if (name) {
		base.name = name;
	}
	const descr = str(cNvPr, 'descr');
	if (descr) {
		base.description = descr;
	}
	const hidden = str(cNvPr, 'hidden');
	if (hidden === '1' || hidden === 'true') {
		base.hidden = true;
	}
	const t = parseXfrm(xfrm);
	if (t.rotation) {
		base.rotation = t.rotation;
	}
	if (t.flipH) {
		base.flipH = true;
	}
	if (t.flipV) {
		base.flipV = true;
	}
	return base;
}

function refColor(
	style: XmlElement | undefined,
	local: string,
): { idx: number; color?: DrawingColor } {
	const ref = kid(style, local);
	if (!ref) {
		return { idx: 0 };
	}
	const idx = local === 'fontRef' ? 1 : (drawingNumber(ref.attrs.idx) ?? 0);
	return { idx, color: findColor(ref) };
}

/** `xdr:sp` and `xdr:cxnSp`. The shape style (`xdr:style`) supplies fallback colors. */
export function buildShape(
	reader: DrawingReader,
	el: XmlElement,
	anchor: DrawingAnchor,
	raw: XmlElement,
	connector: boolean,
): ShapeObject {
	const nv = kid(el, connector ? 'nvCxnSpPr' : 'nvSpPr');
	const spPr = kid(el, 'spPr');
	const shape: ShapeObject = {
		...baseFields(reader, kid(nv, 'cNvPr'), anchor, raw, kid(spPr, 'xfrm')),
		kind: 'shape',
		geometry: 'rect',
	};
	const prst = kid(spPr, 'prstGeom');
	if (prst) {
		shape.geometry = str(prst, 'prst') ?? 'rect';
		const adjust: Record<string, number> = {};
		for (const gd of kids(kid(prst, 'avLst'), 'gd')) {
			const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec((str(gd, 'fmla') ?? '').trim());
			const name = str(gd, 'name');
			if (m && name) {
				adjust[name] = Number(m[1]);
			}
		}
		if (Object.keys(adjust).length > 0) {
			shape.adjust = adjust;
		}
	} else if (kid(spPr, 'custGeom')) {
		shape.geometry = 'custom';
	}
	if (connector) {
		shape.isConnector = true;
	}
	const txBox = str(kid(nv, 'cNvSpPr'), 'txBox');
	if (txBox === '1' || txBox === 'true') {
		shape.isTextBox = true;
	}

	const style = kid(el, 'style');
	const fillRef = refColor(style, 'fillRef');
	const lnRef = refColor(style, 'lnRef');
	const fontRef = refColor(style, 'fontRef');

	let fill: ShapeFill | undefined = findFill(spPr, reader.blipResolver);
	if (!fill && !hasFillElement(spPr) && fillRef.idx > 0 && fillRef.color) {
		fill = { type: 'solid', color: fillRef.color };
	}
	if (fill) {
		shape.fill = fill;
	}

	let line: ShapeLine | undefined = parseLine(kid(spPr, 'ln'));
	if (lnRef.idx > 0 && lnRef.color && !line?.fill) {
		line = { ...line, fill: { type: 'solid', color: lnRef.color } };
		if (line.width === undefined) {
			line.width = THEME_LINE_WIDTHS[Math.min(lnRef.idx, 3)];
		}
	}
	if (line) {
		shape.line = line;
	}

	const text = parseTextBody(kid(el, 'txBody'), fontRef.color);
	if (text && text.paragraphs.some((p) => p.runs.some((r) => r.text !== ''))) {
		shape.text = text;
	}
	return shape;
}

/** `xdr:pic`. Returns `undefined` when the image part cannot be found. */
export async function buildPicture(
	reader: DrawingReader,
	el: XmlElement,
	anchor: DrawingAnchor,
	raw: XmlElement,
): Promise<PictureObject | undefined> {
	const cNvPr = kid(kid(el, 'nvPicPr'), 'cNvPr');
	const spPr = kid(el, 'spPr');
	const blipFill = kid(el, 'blipFill');
	const blip = kid(blipFill, 'blip');
	const relId = relAttr(blip, 'embed') ?? relAttr(blip, 'link');
	const mediaPath = await reader.loadImage(relId);
	if (!mediaPath) {
		return undefined;
	}
	const pic: PictureObject = {
		...baseFields(reader, cNvPr, anchor, raw, kid(spPr, 'xfrm')),
		kind: 'picture',
		mediaPath,
	};
	const src = kid(blipFill, 'srcRect');
	if (src) {
		const crop: NonNullable<PictureObject['crop']> = {};
		for (const side of ['l', 't', 'r', 'b'] as const) {
			const v = drawingNumber(src.attrs[side]);
			if (v) {
				crop[side] = v;
			}
		}
		if (Object.keys(crop).length > 0) {
			pic.crop = crop;
		}
	}
	const line = parseLine(kid(spPr, 'ln'));
	if (line) {
		pic.line = line;
	}
	const link = reader.rels.byId(relAttr(kid(cNvPr, 'hlinkClick'), 'id'));
	if (link) {
		pic.hyperlink = link.target;
	}
	return pic;
}

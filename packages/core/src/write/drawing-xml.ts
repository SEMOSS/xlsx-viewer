/**
 * SpreadsheetDrawing (`xdr:`) markup for regenerated drawing objects:
 * anchors, pictures, shapes, connectors, chart frames and groups.
 */
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import type {
	AnchorPoint,
	ChartObject,
	DrawingAnchor,
	DrawingObject,
	GroupObject,
	PictureObject,
	ShapeObject,
} from '../types/drawing';
import { elements, findDescendantLocal, localName, serializeElement } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { type BlipRelResolver, fillXml, lineXml, textBodyXml } from './drawingml';
import type { SheetGeometry } from './sheet-geometry';
import { flag, tag } from './xml-utils';

export interface Box {
	x: number;
	y: number;
	cx: number;
	cy: number;
}

/** Callbacks into the drawing part writer (relationships, ids, warnings). */
export interface DrawingXmlContext {
	geometry: SheetGeometry;
	imageRel: BlipRelResolver;
	chartRel(obj: ChartObject): string | undefined;
	hyperlinkRel(target: string): string;
	/** A drawing-unique `cNvPr` id, preferring the model id when it is numeric and free. */
	nextId(preferred: string): number;
	/** Original markup can be re-emitted in this part (see `isSelfContained`). */
	reuse(raw: XmlElement): boolean;
	warn(code: string, message: string): void;
}

const R = `xmlns:r="${NS.r}"`;

function int(n: number, min: number, max: number): number {
	return Math.max(min, Math.min(max, Math.round(Number.isFinite(n) ? n : 0)));
}

function pointXml(name: string, p: AnchorPoint): string {
	return (
		`<xdr:${name}><xdr:col>${int(p.col, 0, MAX_COL_INDEX)}</xdr:col>` +
		`<xdr:colOff>${int(p.colOff, 0, 2 ** 31 - 1)}</xdr:colOff>` +
		`<xdr:row>${int(p.row, 0, MAX_ROW_INDEX)}</xdr:row>` +
		`<xdr:rowOff>${int(p.rowOff, 0, 2 ** 31 - 1)}</xdr:rowOff></xdr:${name}>`
	);
}

function extXml(ext: { cx: number; cy: number }): string {
	return `<xdr:ext cx="${int(ext.cx, 0, 2 ** 53)}" cy="${int(ext.cy, 0, 2 ** 53)}"/>`;
}

/** The anchor-position children (`from`/`to`, `ext`, `pos`) of an anchor. */
export function anchorChildrenXml(anchor: DrawingAnchor): string {
	switch (anchor.type) {
		case 'twoCell':
			return pointXml('from', anchor.from) + pointXml('to', anchor.to);
		case 'oneCell':
			return pointXml('from', anchor.from) + extXml(anchor.ext);
		default:
			return `<xdr:pos x="${Math.round(anchor.pos.x)}" y="${Math.round(anchor.pos.y)}"/>${extXml(anchor.ext)}`;
	}
}

/** Wrap object markup in its anchor element. */
export function anchorXml(
	anchor: DrawingAnchor,
	inner: string,
	clientData = '<xdr:clientData/>',
): string {
	const body = anchorChildrenXml(anchor) + inner + clientData;
	switch (anchor.type) {
		case 'twoCell': {
			const editAs =
				anchor.editAs && anchor.editAs !== 'twoCell' ? ` editAs="${anchor.editAs}"` : '';
			return `<xdr:twoCellAnchor${editAs}>${body}</xdr:twoCellAnchor>`;
		}
		case 'oneCell':
			return `<xdr:oneCellAnchor>${body}</xdr:oneCellAnchor>`;
		default:
			return `<xdr:absoluteAnchor>${body}</xdr:absoluteAnchor>`;
	}
}

function xfrmXml(
	name: string,
	obj: DrawingObject,
	box: Box,
	child?: { off: string; ext: string },
): string {
	const attrs = {
		rot: obj.rotation ? Math.round(obj.rotation * 60000) : undefined,
		flipH: flag(obj.flipH),
		flipV: flag(obj.flipV),
	};
	const inner =
		`<a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/>` +
		`<a:ext cx="${Math.max(0, Math.round(box.cx))}" cy="${Math.max(0, Math.round(box.cy))}"/>` +
		(child ? child.off + child.ext : '');
	return tag(name, attrs, inner);
}

function cNvPrXml(obj: DrawingObject, id: number, fallbackName: string, inner = ''): string {
	return tag(
		'xdr:cNvPr',
		{
			id,
			name: obj.name ?? `${fallbackName} ${Math.max(1, id - 1)}`,
			descr: obj.description,
			hidden: flag(obj.hidden),
		},
		inner,
	);
}

/** The `xdr:style` element of the original markup, if any (kept for regenerated shapes). */
function rawStyle(obj: DrawingObject): string | undefined {
	const style = obj.raw ? findDescendantLocal(obj.raw, 'style') : undefined;
	return style && style.name.endsWith(':style') ? serializeElement(style) : undefined;
}

function schemeRef(name: string, idx: string | number, color: string, shade = false): string {
	const inner = shade
		? `<a:schemeClr val="${color}"><a:shade val="50000"/></a:schemeClr>`
		: `<a:schemeClr val="${color}"/>`;
	return `<a:${name} idx="${idx}">${inner}</a:${name}>`;
}

/** Excel's default style references for new shapes, text boxes and connectors. */
function defaultStyle(kind: 'shape' | 'textBox' | 'connector'): string {
	if (kind === 'textBox') {
		return `<xdr:style>${schemeRef('lnRef', 0, 'accent1')}${schemeRef('fillRef', 0, 'accent1')}${schemeRef('effectRef', 0, 'accent1')}${schemeRef('fontRef', 'minor', 'dk1')}</xdr:style>`;
	}
	if (kind === 'connector') {
		return `<xdr:style>${schemeRef('lnRef', 1, 'accent1')}${schemeRef('fillRef', 0, 'accent1')}${schemeRef('effectRef', 0, 'accent1')}${schemeRef('fontRef', 'minor', 'tx1')}</xdr:style>`;
	}
	return `<xdr:style>${schemeRef('lnRef', 2, 'accent1', true)}${schemeRef('fillRef', 1, 'accent1')}${schemeRef('effectRef', 0, 'accent1')}${schemeRef('fontRef', 'minor', 'lt1')}</xdr:style>`;
}

function geometryXml(obj: ShapeObject): string {
	if (obj.geometry === 'custom' || !/^[A-Za-z0-9]+$/.test(obj.geometry)) {
		const cust = obj.raw ? findDescendantLocal(obj.raw, 'custGeom') : undefined;
		return cust ? serializeElement(cust) : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
	}
	const gd = Object.entries(obj.adjust ?? {})
		.filter(([name, v]) => /^[A-Za-z0-9]+$/.test(name) && Number.isFinite(v))
		.map(([name, v]) => `<a:gd name="${name}" fmla="val ${Math.round(v)}"/>`)
		.join('');
	return `<a:prstGeom prst="${obj.geometry}">${gd ? `<a:avLst>${gd}</a:avLst>` : '<a:avLst/>'}</a:prstGeom>`;
}

const TEXT_BOX_FILL = '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>';
const TEXT_BOX_LINE =
	'<a:ln w="9525" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>';

function shapeXml(obj: ShapeObject, ctx: DrawingXmlContext, box: Box): string {
	const id = ctx.nextId(obj.id);
	const fresh = !obj.raw;
	let fill = obj.fill ? fillXml(obj.fill, ctx.imageRel) : '';
	let line = obj.line ? lineXml(obj.line, ctx.imageRel) : '';
	if (obj.isTextBox && fresh) {
		fill ||= TEXT_BOX_FILL;
		line ||= TEXT_BOX_LINE;
	}
	const spPr = `<xdr:spPr>${xfrmXml('a:xfrm', obj, box)}${geometryXml(obj)}${fill}${line}</xdr:spPr>`;
	if (obj.isConnector) {
		const nv = `<xdr:nvCxnSpPr>${cNvPrXml(obj, id, 'Straight Connector')}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>`;
		return `<xdr:cxnSp macro="">${nv}${spPr}${rawStyle(obj) ?? defaultStyle('connector')}</xdr:cxnSp>`;
	}
	const nv = `<xdr:nvSpPr>${cNvPrXml(obj, id, obj.isTextBox ? 'TextBox' : 'Shape')}${tag(
		'xdr:cNvSpPr',
		{
			txBox: flag(obj.isTextBox),
		},
	)}</xdr:nvSpPr>`;
	const style = rawStyle(obj) ?? defaultStyle(obj.isTextBox ? 'textBox' : 'shape');
	const text = obj.text || obj.isTextBox ? textBodyXml(obj.text) : '';
	return `<xdr:sp macro="" textlink="">${nv}${spPr}${style}${text}</xdr:sp>`;
}

function pictureXml(obj: PictureObject, ctx: DrawingXmlContext, box: Box): string | undefined {
	const rel = ctx.imageRel(obj.mediaPath);
	if (!rel) {
		ctx.warn('missing-media', `Picture "${obj.name ?? obj.id}" has no image data and was skipped`);
		return undefined;
	}
	const id = ctx.nextId(obj.id);
	const link = obj.hyperlink
		? `<a:hlinkClick ${R} r:id="${ctx.hyperlinkRel(obj.hyperlink)}"/>`
		: '';
	const crop = obj.crop
		? tag('a:srcRect', {
				l: obj.crop.l ? Math.round(obj.crop.l) : undefined,
				t: obj.crop.t ? Math.round(obj.crop.t) : undefined,
				r: obj.crop.r ? Math.round(obj.crop.r) : undefined,
				b: obj.crop.b ? Math.round(obj.crop.b) : undefined,
			})
		: '';
	return (
		`<xdr:pic><xdr:nvPicPr>${cNvPrXml(obj, id, 'Picture', link)}` +
		'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
		`<xdr:blipFill><a:blip ${R} r:embed="${rel}"/>${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
		`<xdr:spPr>${xfrmXml('a:xfrm', obj, box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
		`${obj.line ? lineXml(obj.line) : ''}</xdr:spPr></xdr:pic>`
	);
}

function chartFrameXml(obj: ChartObject, ctx: DrawingXmlContext, box: Box): string | undefined {
	const rel = ctx.chartRel(obj);
	if (!rel) {
		return undefined;
	}
	const id = ctx.nextId(obj.id);
	return (
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>${cNvPrXml(obj, id, 'Chart')}` +
		'<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
		xfrmXml('xdr:xfrm', obj, box) +
		'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
		`<c:chart xmlns:c="${NS.c}" ${R} r:id="${rel}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
	);
}

/** The shape element itself when `raw` is an anchor (top-level raw used inside a group). */
function innerOfAnchor(raw: XmlElement): XmlElement | undefined {
	const skip = new Set(['from', 'to', 'ext', 'pos', 'clientData']);
	return elements(raw).find((c) => !skip.has(localName(c.name)));
}

function groupXml(obj: GroupObject, ctx: DrawingXmlContext, box: Box): string {
	const id = ctx.nextId(obj.id);
	const chOff = obj.childOffset ?? { x: box.x, y: box.y };
	const chExt = obj.childExtent ?? { cx: box.cx, cy: box.cy };
	const kids = obj.children
		.map((c) => childXml(c, ctx))
		.filter((s): s is string => !!s)
		.join('');
	const xfrm = xfrmXml('a:xfrm', obj, box, {
		off: `<a:chOff x="${Math.round(chOff.x)}" y="${Math.round(chOff.y)}"/>`,
		ext: `<a:chExt cx="${Math.round(chExt.cx)}" cy="${Math.round(chExt.cy)}"/>`,
	});
	return (
		`<xdr:grpSp><xdr:nvGrpSpPr>${cNvPrXml(obj, id, 'Group')}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
		`<xdr:grpSpPr>${xfrm}</xdr:grpSpPr>${kids}</xdr:grpSp>`
	);
}

/** Markup of an object without its anchor; `undefined` when it cannot be written. */
export function objectXml(
	obj: DrawingObject,
	ctx: DrawingXmlContext,
	box: Box,
): string | undefined {
	switch (obj.kind) {
		case 'picture':
			return pictureXml(obj, ctx, box);
		case 'chart':
			return chartFrameXml(obj, ctx, box);
		case 'shape':
			return shapeXml(obj, ctx, box);
		case 'group':
			return groupXml(obj, ctx, box);
		default: {
			const inner =
				obj.raw && localName(obj.raw.name).endsWith('Anchor') ? innerOfAnchor(obj.raw) : obj.raw;
			if (inner && ctx.reuse(inner)) {
				return serializeElement(inner);
			}
			ctx.warn(
				'drawing-object-dropped',
				`"${obj.name ?? obj.id}" (${obj.label ?? 'object'}) cannot be saved here`,
			);
			return undefined;
		}
	}
}

/** A group child: verbatim when untouched, regenerated otherwise (absolute child coordinates). */
function childXml(obj: DrawingObject, ctx: DrawingXmlContext): string | undefined {
	if (obj.raw && obj.dirty === false && ctx.reuse(obj.raw)) {
		const inner = localName(obj.raw.name).endsWith('Anchor') ? innerOfAnchor(obj.raw) : obj.raw;
		return inner ? serializeElement(inner) : undefined;
	}
	const box =
		obj.anchor.type === 'absolute'
			? { ...obj.anchor.pos, ...obj.anchor.ext }
			: ctx.geometry.anchorBox(obj.anchor);
	return objectXml(obj, ctx, box);
}

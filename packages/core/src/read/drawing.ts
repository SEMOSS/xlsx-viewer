import type { AnchorPoint, DrawingAnchor, DrawingObject } from '../types/drawing';
import { isElement, textContent } from '../xml/dom';
import { parseXml } from '../xml/parser';
import type { XmlElement } from '../xml/types';
import type { ReadContext } from './context';
import { buildObject, chooseAlternate, isUnderstoodObject } from './drawing-objects';
import { DrawingReader } from './drawing-reader';
import { coordinate } from './drawingml';
import { readPartText, relLookup } from './part-io';
import { kid, lname, str } from './xml-utils';

/** Drawing parts (`xl/drawings/drawingN.xml`, `xdr:wsDr`) -> drawing objects. */

const ANCHORS: ReadonlySet<string> = new Set(['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']);
const OBJECTS: ReadonlySet<string> = new Set([
	'sp',
	'cxnSp',
	'pic',
	'grpSp',
	'graphicFrame',
	'AlternateContent',
	'contentPart',
]);

function anchorPoint(el: XmlElement | undefined): AnchorPoint {
	const n = (local: string): number => {
		const v = coordinate(textContent(kid(el, local)));
		return v === undefined ? 0 : Math.max(0, Math.round(v));
	};
	return { col: n('col'), colOff: n('colOff'), row: n('row'), rowOff: n('rowOff') };
}

function extent(el: XmlElement | undefined): { cx: number; cy: number } {
	return { cx: coordinate(el?.attrs.cx) ?? 0, cy: coordinate(el?.attrs.cy) ?? 0 };
}

/** The anchor geometry of `xdr:twoCellAnchor` / `xdr:oneCellAnchor` / `xdr:absoluteAnchor`. */
export function parseAnchor(el: XmlElement): DrawingAnchor | undefined {
	switch (lname(el)) {
		case 'twoCellAnchor': {
			const anchor: DrawingAnchor = {
				type: 'twoCell',
				from: anchorPoint(kid(el, 'from')),
				to: anchorPoint(kid(el, 'to')),
			};
			const editAs = str(el, 'editAs');
			if (editAs === 'twoCell' || editAs === 'oneCell' || editAs === 'absolute') {
				anchor.editAs = editAs;
			}
			return anchor;
		}
		case 'oneCellAnchor':
			return { type: 'oneCell', from: anchorPoint(kid(el, 'from')), ext: extent(kid(el, 'ext')) };
		case 'absoluteAnchor': {
			const pos = kid(el, 'pos');
			return {
				type: 'absolute',
				pos: { x: coordinate(pos?.attrs.x) ?? 0, y: coordinate(pos?.attrs.y) ?? 0 },
				ext: extent(kid(el, 'ext')),
			};
		}
		default:
			return undefined;
	}
}

function objectElement(anchorEl: XmlElement): XmlElement | undefined {
	for (const c of anchorEl.children) {
		if (isElement(c) && OBJECTS.has(lname(c))) {
			return c;
		}
	}
	return undefined;
}

function isUnderstoodAnchor(el: XmlElement): boolean {
	return ANCHORS.has(lname(el)) && isUnderstoodObject(objectElement(el));
}

async function readAnchor(
	reader: DrawingReader,
	anchorEl: XmlElement,
	raw: XmlElement,
): Promise<DrawingObject | undefined> {
	const anchor = parseAnchor(anchorEl);
	const obj = anchor ? objectElement(anchorEl) : undefined;
	if (!anchor || !obj) {
		return undefined;
	}
	return buildObject(reader, obj, anchor, raw);
}

/** Read a parsed `xdr:wsDr` root. */
export async function readDrawingRoot(
	reader: DrawingReader,
	root: XmlElement,
): Promise<DrawingObject[]> {
	const out: DrawingObject[] = [];
	for (const c of root.children) {
		if (!isElement(c)) {
			continue;
		}
		const name = lname(c);
		let anchorEl: XmlElement | undefined;
		if (ANCHORS.has(name)) {
			anchorEl = c;
		} else if (name === 'AlternateContent') {
			const chosen = chooseAlternate(reader, c, isUnderstoodAnchor);
			anchorEl = chosen && ANCHORS.has(lname(chosen)) ? chosen : undefined;
		}
		if (!anchorEl) {
			continue;
		}
		try {
			const obj = await readAnchor(reader, anchorEl, c);
			if (obj) {
				out.push(obj);
			}
		} catch (e) {
			reader.ctx.unreadable(reader.partPath, e, 'drawing object in', reader.sheetName);
		}
	}
	await reader.settle();
	return out;
}

/** Read a drawing part with its relationships. Missing parts yield no objects and a warning. */
export async function readDrawingPart(
	ctx: ReadContext,
	partPath: string,
	sheetName: string,
): Promise<DrawingObject[]> {
	const xml = await readPartText(ctx.pkg, partPath);
	if (xml === undefined) {
		ctx.warn('missing-part', `Drawing part ${partPath} is missing`, {
			part: partPath,
			sheet: sheetName,
		});
		return [];
	}
	const rels = relLookup(partPath, await ctx.pkg.rels(partPath), ctx.pkg);
	const reader = new DrawingReader(ctx, partPath, rels, sheetName);
	return readDrawingRoot(reader, parseXml(xml).root);
}

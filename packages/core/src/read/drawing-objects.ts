import { parseChartXml } from '../charts';
import type {
	ChartObject,
	DrawingAnchor,
	DrawingObject,
	GroupObject,
	UnknownDrawingObject,
} from '../types/drawing';
import { findDescendantLocal, isElement } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import type { DrawingReader } from './drawing-reader';
import { baseFields, buildPicture, buildShape } from './drawing-shapes';
import { parseXfrm } from './drawingml';
import { readPartText } from './part-io';
import { firstElement, kid, kids, lname, relAttr, str } from './xml-utils';

/** Object dispatch for drawing parts: groups, graphic frames (charts) and alternate content. */

const CHART_URI = /\/drawingml\/(?:2006\/)?chart$/;

function graphicUri(frame: XmlElement): string {
	return str(kid(kid(frame, 'graphic'), 'graphicData'), 'uri') ?? '';
}

/** Objects this reader models fully (used to pick a branch of `mc:AlternateContent`). */
export function isUnderstoodObject(el: XmlElement | undefined): boolean {
	if (!el) {
		return false;
	}
	switch (lname(el)) {
		case 'sp':
		case 'cxnSp':
		case 'pic':
		case 'grpSp':
			return true;
		case 'graphicFrame':
			return CHART_URI.test(graphicUri(el));
		default:
			return false;
	}
}

/**
 * Pick the branch of `mc:AlternateContent`: the first `mc:Choice` whose
 * content passes `understood`, else `mc:Fallback`, else the first choice.
 * Warns about Excel 2016+ charts (chartex) found in a choice.
 */
export function chooseAlternate(
	reader: DrawingReader,
	ac: XmlElement,
	understood: (el: XmlElement) => boolean,
): XmlElement | undefined {
	const choices = kids(ac, 'Choice');
	for (const choice of choices) {
		const data = findDescendantLocal(choice, 'graphicData');
		if (data && (str(data, 'uri') ?? '').includes('chartex')) {
			reader.warn(
				'unsupported-chart',
				'The sheet contains an Excel 2016+ chart type (waterfall, treemap, funnel, ...) that is not rendered',
			);
		}
	}
	for (const choice of choices) {
		const content = firstElement(choice);
		if (content && understood(content)) {
			return content;
		}
	}
	return firstElement(kid(ac, 'Fallback')) ?? firstElement(choices[0]);
}

function unknownObject(
	reader: DrawingReader,
	cNvPr: XmlElement | undefined,
	anchor: DrawingAnchor,
	raw: XmlElement,
	label: string,
	xfrm?: XmlElement,
): UnknownDrawingObject {
	return { ...baseFields(reader, cNvPr, anchor, raw, xfrm), kind: 'unknown', label };
}

async function buildGraphicFrame(
	reader: DrawingReader,
	el: XmlElement,
	anchor: DrawingAnchor,
	raw: XmlElement,
): Promise<DrawingObject> {
	const cNvPr = kid(kid(el, 'nvGraphicFramePr'), 'cNvPr');
	const xfrm = kid(el, 'xfrm');
	const uri = graphicUri(el);
	if (CHART_URI.test(uri)) {
		const chartEl = kid(kid(kid(el, 'graphic'), 'graphicData'), 'chart');
		const chartPath = reader.rels.path(relAttr(chartEl, 'id'));
		const xml = chartPath ? await readPartText(reader.ctx.pkg, chartPath) : undefined;
		if (!chartPath || xml === undefined) {
			reader.warn('missing-part', 'A chart part referenced by the drawing is missing');
			return unknownObject(reader, cNvPr, anchor, raw, 'Chart (missing part)', xfrm);
		}
		let chart: ChartObject['chart'];
		try {
			chart = parseChartXml(xml, { partPath: chartPath });
		} catch (e) {
			reader.ctx.unreadable(chartPath, e, 'chart', reader.sheetName);
			chart = { plots: [], axes: [], rawXml: xml };
		}
		if (chart.plots.some((p) => p.type === 'unknown')) {
			reader.warn(
				'unsupported-chart',
				'A chart uses a chart type that is not supported',
				chartPath,
			);
		}
		return { ...baseFields(reader, cNvPr, anchor, raw, xfrm), kind: 'chart', chart, chartPath };
	}
	let label = 'Embedded object';
	if (uri.includes('chartex')) {
		label = 'Chart (Excel 2016+ chart type)';
		reader.warn(
			'unsupported-chart',
			'The sheet contains an Excel 2016+ chart type (waterfall, treemap, funnel, ...) that is not rendered',
		);
	} else if (uri.endsWith('/diagram')) {
		label = 'SmartArt';
	} else if (uri.includes('timeslicer')) {
		label = 'Timeline';
		reader.warn('slicer', 'Timelines are shown as placeholders and cannot be used');
	} else if (uri.includes('slicer')) {
		label = 'Slicer';
		reader.warn('slicer', 'Slicers are shown as placeholders and cannot be used');
	} else if (uri.includes('/ole') || uri.includes('presentationml')) {
		label = 'OLE object';
		reader.warn('ole-object', 'Embedded OLE objects are shown as placeholders');
	}
	return unknownObject(reader, cNvPr, anchor, raw, label, xfrm);
}

/** Transform element of a group child (where its position in child coordinates lives). */
function childXfrm(el: XmlElement): XmlElement | undefined {
	switch (lname(el)) {
		case 'grpSp':
			return kid(kid(el, 'grpSpPr'), 'xfrm');
		case 'graphicFrame':
			return kid(el, 'xfrm');
		default:
			return kid(kid(el, 'spPr'), 'xfrm');
	}
}

async function buildGroup(
	reader: DrawingReader,
	el: XmlElement,
	anchor: DrawingAnchor,
	raw: XmlElement,
): Promise<GroupObject> {
	const grpSpPr = kid(el, 'grpSpPr');
	const xfrmEl = kid(grpSpPr, 'xfrm');
	const group: GroupObject = {
		...baseFields(reader, kid(kid(el, 'nvGrpSpPr'), 'cNvPr'), anchor, raw, xfrmEl),
		kind: 'group',
		children: [],
	};
	const xfrm = parseXfrm(xfrmEl);
	if (xfrm.chOff) {
		group.childOffset = xfrm.chOff;
	}
	if (xfrm.chExt) {
		group.childExtent = xfrm.chExt;
	}
	for (const c of el.children) {
		if (!isElement(c)) {
			continue;
		}
		const name = lname(c);
		if (name === 'nvGrpSpPr' || name === 'grpSpPr') {
			continue;
		}
		let target: XmlElement | undefined = c;
		if (name === 'AlternateContent') {
			target = chooseAlternate(reader, c, isUnderstoodObject);
		}
		if (!target) {
			continue;
		}
		const t = parseXfrm(childXfrm(target));
		const childAnchor: DrawingAnchor = {
			type: 'absolute',
			pos: t.off ?? { x: 0, y: 0 },
			ext: t.ext ?? { cx: 0, cy: 0 },
		};
		const child = await buildObject(reader, target, childAnchor, c);
		if (child) {
			group.children.push(child);
		}
	}
	return group;
}

/**
 * Build the model object for a drawing element. `raw` is the element kept
 * for round-trip: the top-level anchor (or the `mc:AlternateContent` that
 * wraps it) for top-level objects, the child element itself inside groups.
 */
export async function buildObject(
	reader: DrawingReader,
	el: XmlElement,
	anchor: DrawingAnchor,
	raw: XmlElement,
): Promise<DrawingObject | undefined> {
	switch (lname(el)) {
		case 'sp':
			return buildShape(reader, el, anchor, raw, false);
		case 'cxnSp':
			return buildShape(reader, el, anchor, raw, true);
		case 'pic': {
			const pic = await buildPicture(reader, el, anchor, raw);
			if (pic) {
				return pic;
			}
			reader.warn('missing-part', 'A picture refers to an image that is missing from the file');
			return unknownObject(
				reader,
				kid(kid(el, 'nvPicPr'), 'cNvPr'),
				anchor,
				raw,
				'Picture (image missing)',
			);
		}
		case 'grpSp':
			return buildGroup(reader, el, anchor, raw);
		case 'graphicFrame':
			return buildGraphicFrame(reader, el, anchor, raw);
		case 'AlternateContent': {
			const chosen = chooseAlternate(reader, el, isUnderstoodObject);
			return chosen ? buildObject(reader, chosen, anchor, raw) : undefined;
		}
		case 'contentPart':
			return unknownObject(reader, undefined, anchor, raw, 'Ink');
		default:
			return undefined;
	}
}

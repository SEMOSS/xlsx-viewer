import type { ChartObject, DrawingAnchor } from '../types/drawing';
import type { Worksheet } from '../types/sheet';
import type { XmlElement } from '../xml/types';
import { attempt, type ReadContext } from './context';
import { readDrawingPart } from './drawing';
import type { RelLookup } from './part-io';
import { applyPrintSettings, applySheetPr, parseSheetProtection } from './sheet-features';
import { parseSheetView } from './sheet-view';
import { kid, str } from './xml-utils';

/** Chart size used when a chartsheet drawing does not state an extent (EMU, about 9.5 x 6.9 inches). */
const DEFAULT_CHARTSHEET_EXTENT = { cx: 8_686_800, cy: 6_291_943 };

function sheetExtent(anchor: DrawingAnchor): { cx: number; cy: number } {
	if (anchor.type === 'absolute' || anchor.type === 'oneCell') {
		if (anchor.ext.cx > 0 && anchor.ext.cy > 0) {
			return anchor.ext;
		}
	}
	return DEFAULT_CHARTSHEET_EXTENT;
}

/**
 * A chartsheet (`xl/chartsheets/sheetN.xml`): a sheet that is one chart.
 * The chart becomes the only drawing object, anchored absolutely at the
 * sheet origin with the drawing's extent.
 */
export async function readChartsheet(
	ctx: ReadContext,
	sheet: Worksheet,
	root: XmlElement,
	rels: RelLookup,
): Promise<void> {
	applySheetPr(sheet, kid(root, 'sheetPr'));
	sheet.view = parseSheetView(kid(root, 'sheetViews'));
	const protection = parseSheetProtection(kid(root, 'sheetProtection'));
	if (protection) {
		sheet.protection = protection;
	}
	applyPrintSettings(sheet, root);
	const drawingPath = rels.path(str(kid(root, 'drawing'), 'r:id'));
	if (!drawingPath) {
		ctx.warn('missing-part', `Chartsheet "${sheet.name}" has no chart drawing`, {
			sheet: sheet.name,
		});
		return;
	}
	if (sheet.source) {
		sheet.source.drawingPath = drawingPath;
	}
	if (ctx.options.skipDrawings) {
		return;
	}
	const objects = await attempt(
		ctx,
		drawingPath,
		'drawing',
		() => readDrawingPart(ctx, drawingPath, sheet.name),
		[],
		sheet.name,
	);
	const chart = objects.find((o): o is ChartObject => o.kind === 'chart');
	if (!chart) {
		ctx.warn('unsupported-chart', `Chartsheet "${sheet.name}" does not contain a readable chart`, {
			part: drawingPath,
			sheet: sheet.name,
		});
		return;
	}
	chart.anchor = { type: 'absolute', pos: { x: 0, y: 0 }, ext: sheetExtent(chart.anchor) };
	sheet.drawings = [chart];
}

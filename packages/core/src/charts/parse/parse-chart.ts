/**
 * DrawingML chart part (`c:chartSpace`) -> {@link ChartModel}.
 *
 * The parser is tolerant: anything it does not model stays in `rawXml`, which
 * the writer re-emits verbatim while the chart is unedited (`dirty === false`)
 * and merges back into regenerated XML after edits.
 */
import type { ChartModel, ChartPlot, ChartPlotType } from '../../types/chart';
import type { XlsxWarning } from '../../types/warnings';
import { findDescendantLocal, isElement, localName } from '../../xml/dom';
import { parseXmlRoot } from '../../xml/parser';
import type { XmlElement } from '../../xml/types';
import { boolOf, kid, kids, numOf, readShapeProps, valOf } from '../xml/drawingml-read';
import { readManualLayout, readRichText, readTxPr } from '../xml/text-read';
import { isAxisElement, readAxis, readLegend, readTitle } from './parse-axis';
import { readDataLabels, readSeries } from './parse-series';

export interface ChartParseContext {
	/** Package path of the chart part, for warnings. */
	partPath?: string;
	/** Receives approximations and unsupported content found while parsing. */
	onWarning?: (warning: XlsxWarning) => void;
}

/** Plot container local name -> model type and 3D flag. */
const PLOT_ELEMENTS: Record<string, { type: ChartPlotType; is3D?: boolean }> = {
	barChart: { type: 'bar' },
	bar3DChart: { type: 'bar', is3D: true },
	lineChart: { type: 'line' },
	line3DChart: { type: 'line', is3D: true },
	pieChart: { type: 'pie' },
	pie3DChart: { type: 'pie', is3D: true },
	doughnutChart: { type: 'doughnut' },
	areaChart: { type: 'area' },
	area3DChart: { type: 'area', is3D: true },
	scatterChart: { type: 'scatter' },
	radarChart: { type: 'radar' },
	bubbleChart: { type: 'bubble' },
	stockChart: { type: 'stock' },
	surfaceChart: { type: 'surface' },
	surface3DChart: { type: 'surface', is3D: true },
	ofPieChart: { type: 'ofPie' },
};

export function isPlotElement(local: string): boolean {
	return local in PLOT_ELEMENTS;
}

const GROUPINGS = new Set(['clustered', 'stacked', 'percentStacked', 'standard']);
const SCATTER_STYLES = new Set(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker', 'none']);
const RADAR_STYLES = new Set(['standard', 'marker', 'filled']);

function readPlot(el: XmlElement, local: string): ChartPlot {
	const info = PLOT_ELEMENTS[local];
	const plot: ChartPlot = {
		type: info.type,
		series: kids(el, 'ser').map((s, i) => readSeries(s, i)),
		axisIds: kids(el, 'axId')
			.map((a) => Number(a.attrs.val))
			.filter((n) => Number.isFinite(n)),
	};
	if (info.is3D) {
		plot.is3D = true;
	}
	const barDir = valOf(el, 'barDir');
	if (barDir === 'bar' || barDir === 'col') {
		plot.barDirection = barDir;
	} else if (info.type === 'bar') {
		plot.barDirection = 'col';
	}
	const grouping = valOf(el, 'grouping');
	if (grouping && GROUPINGS.has(grouping)) {
		plot.grouping = grouping as ChartPlot['grouping'];
	}
	const vary = boolOf(el, 'varyColors');
	if (vary !== undefined) {
		plot.varyColors = vary;
	}
	const numbers: Array<
		[string, 'gapWidth' | 'overlap' | 'holeSize' | 'firstSliceAngle' | 'bubbleScale']
	> = [
		['gapWidth', 'gapWidth'],
		['overlap', 'overlap'],
		['holeSize', 'holeSize'],
		['firstSliceAng', 'firstSliceAngle'],
		['bubbleScale', 'bubbleScale'],
	];
	for (const [tag, key] of numbers) {
		const v = numOf(el, tag);
		if (v !== undefined) {
			plot[key] = v;
		}
	}
	const scatterStyle = valOf(el, 'scatterStyle');
	if (scatterStyle && SCATTER_STYLES.has(scatterStyle)) {
		plot.scatterStyle = scatterStyle as ChartPlot['scatterStyle'];
	}
	const radarStyle = valOf(el, 'radarStyle');
	if (radarStyle && RADAR_STYLES.has(radarStyle)) {
		plot.radarStyle = radarStyle as ChartPlot['radarStyle'];
	}
	const labels = readDataLabels(kid(el, 'dLbls'));
	if (labels) {
		plot.dataLabels = labels;
	}
	const marker = boolOf(el, 'marker');
	if (marker !== undefined) {
		plot.marker = marker;
	}
	if (kid(el, 'hiLowLines')) {
		plot.hiLowLines = true;
	}
	if (kid(el, 'upDownBars')) {
		plot.upDownBars = true;
	}
	const sizeRepresents = valOf(el, 'sizeRepresents');
	if (sizeRepresents === 'area' || sizeRepresents === 'w') {
		plot.bubbleSizeRepresents = sizeRepresents;
	}
	return plot;
}

/** `c:style`, or the richer `c14:style` inside `mc:AlternateContent`. */
function readStyle(space: XmlElement): number | undefined {
	for (const c of space.children) {
		if (!isElement(c)) {
			continue;
		}
		const local = localName(c.name);
		if (local === 'style') {
			const n = Number(c.attrs.val);
			if (Number.isFinite(n)) {
				return n;
			}
		} else if (local === 'AlternateContent') {
			for (const branch of [kid(c, 'Choice'), kid(c, 'Fallback')]) {
				const style = kid(branch, 'style');
				const n = Number(style?.attrs.val);
				if (style && Number.isFinite(n)) {
					return n;
				}
			}
		}
	}
	return undefined;
}

const APPROXIMATED: Partial<Record<ChartPlotType, string>> = {
	surface: 'Surface charts are drawn as line charts.',
	ofPie: 'Pie-of-pie and bar-of-pie charts are drawn as plain pies.',
};

function readChartSpace(space: XmlElement, xml: string, ctx: ChartParseContext): ChartModel {
	const model: ChartModel = { plots: [], axes: [], rawXml: xml, dirty: false };
	const chart = kid(space, 'chart');
	const plotArea = kid(chart, 'plotArea');
	for (const c of plotArea?.children ?? []) {
		if (!isElement(c)) {
			continue;
		}
		const local = localName(c.name);
		if (isPlotElement(local)) {
			model.plots.push(readPlot(c, local));
		} else if (isAxisElement(local)) {
			model.axes.push(readAxis(c, local));
		}
	}
	const title = readTitle(kid(chart, 'title'));
	if (title) {
		model.title = title;
	}
	const autoTitleDeleted = boolOf(chart, 'autoTitleDeleted');
	if (autoTitleDeleted !== undefined) {
		model.autoTitleDeleted = autoTitleDeleted;
	}
	const legend = readLegend(kid(chart, 'legend'));
	if (legend) {
		model.legend = legend;
	}
	const visOnly = boolOf(chart, 'plotVisOnly');
	if (visOnly !== undefined) {
		model.plotVisibleOnly = visOnly;
	}
	const blanks = valOf(chart, 'dispBlanksAs');
	if (blanks === 'gap' || blanks === 'zero' || blanks === 'span') {
		model.displayBlanksAs = blanks;
	}
	const style = readStyle(space);
	if (style !== undefined) {
		model.style = style;
	}
	// The schema default for an omitted c:roundedCorners is true, and Excel honours it.
	model.roundedCorners = boolOf(space, 'roundedCorners') ?? true;
	const spaceSp = readShapeProps(kid(space, 'spPr'));
	if (spaceSp.fill) {
		model.chartFill = spaceSp.fill;
	}
	if (spaceSp.line) {
		model.chartLine = spaceSp.line;
	}
	const plotSp = readShapeProps(kid(plotArea, 'spPr'));
	if (plotSp.fill) {
		model.plotFill = plotSp.fill;
	}
	if (plotSp.line) {
		model.plotLine = plotSp.line;
	}
	const plotLayout = readManualLayout(kid(plotArea, 'layout'));
	if (plotLayout) {
		model.plotAreaLayout = plotLayout;
	}
	const { style: textStyle } = readTxPr(kid(space, 'txPr'));
	if (textStyle) {
		model.textStyle = textStyle;
	}
	for (const plot of model.plots) {
		const note =
			APPROXIMATED[plot.type] ?? (plot.is3D ? '3-D charts are drawn in 2-D.' : undefined);
		if (note) {
			ctx.onWarning?.({ code: 'chart-type-approximated', message: note, part: ctx.partPath });
		}
	}
	return model;
}

/** Chart parts of the 2014 `cx:` (chartex) family: keep the XML, expose only the title. */
function readChartEx(space: XmlElement, xml: string, ctx: ChartParseContext): ChartModel {
	const model: ChartModel = { plots: [], axes: [], rawXml: xml, dirty: false };
	const title = findDescendantLocal(space, 'title');
	const rich = title ? findDescendantLocal(title, 'rich') : undefined;
	if (rich) {
		const r = readRichText(rich);
		if (r.text) {
			model.title = { text: r.text, paragraphs: r.paragraphs };
		}
	}
	const layout = findDescendantLocal(space, 'series')?.attrs.layoutId;
	ctx.onWarning?.({
		code: 'chart-unsupported',
		message: `Chart type "${layout ?? 'chartex'}" is not supported; the chart is preserved but not drawn.`,
		part: ctx.partPath,
	});
	return model;
}

/** Parse a chart part (`c:chartSpace`). Unknown content is kept in `rawXml` for round-trip. */
export function parseChartXml(xml: string, ctx: ChartParseContext = {}): ChartModel {
	let root: XmlElement;
	try {
		root = parseXmlRoot(xml);
	} catch (err) {
		ctx.onWarning?.({
			code: 'chart-unreadable',
			message: `Chart part could not be parsed: ${(err as Error).message}`,
			part: ctx.partPath,
		});
		return { plots: [], axes: [], rawXml: xml, dirty: false };
	}
	if (localName(root.name) !== 'chartSpace') {
		return { plots: [], axes: [], rawXml: xml, dirty: false };
	}
	const prefix = root.name.includes(':') ? root.name.slice(0, root.name.indexOf(':')) : '';
	const ns = root.attrs[prefix ? `xmlns:${prefix}` : 'xmlns'] ?? '';
	if (ns.includes('chartex') || (!kid(root, 'chart') && kid(root, 'chartData'))) {
		return readChartEx(root, xml, ctx);
	}
	return readChartSpace(root, xml, ctx);
}

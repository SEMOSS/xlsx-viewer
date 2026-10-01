/**
 * {@link ChartModel} -> DrawingML chart part. Unedited charts are returned
 * verbatim; edited or new charts are regenerated in schema order, with markup
 * the model does not cover merged back from the original XML.
 */
import type { ChartModel } from '../../types/chart';
import { el, isElement, localName } from '../../xml/dom';
import { NS } from '../../xml/namespaces';
import { parseXmlRoot } from '../../xml/parser';
import { serializeXml } from '../../xml/serializer';
import type { XmlElement } from '../../xml/types';
import { valEl, writeLayout, writeShapeProps, writeTextProps } from '../xml/drawingml-write';
import { mergeRawChildren, rawChild, rawChildByKey, rawChildren } from './merge-raw';
import { CHART_ORDER } from './schema';
import { writeAxis, writeLegend, writeTitle } from './write-axis';
import { type PlannedPlot, planPlots } from './plan-plots';
import { writeDataLabels, writeSeries } from './write-series';

function writePlot(planned: PlannedPlot, raw: XmlElement | undefined): XmlElement {
	const { plot, container } = planned;
	const v = (name: string, value: string | number | boolean | undefined) =>
		value === undefined ? undefined : valEl(name, value);
	const series = plot.series.map((s) =>
		writeSeries(s, container, rawChildByKey(raw, 'ser', 'idx', s.idx)),
	);
	const labels = container.startsWith('surface')
		? undefined
		: writeDataLabels(plot.dataLabels, rawChild(raw, 'dLbls'));
	const axIds = planned.axisIds.map((id) => valEl('c:axId', id));
	const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
	const vary = v('c:varyColors', plot.varyColors);
	let children: Array<XmlElement | undefined>;
	switch (container) {
		case 'barChart':
		case 'bar3DChart':
			children = [
				valEl('c:barDir', plot.barDirection ?? 'col'),
				valEl('c:grouping', plot.grouping ?? 'clustered'),
				vary,
				...series,
				labels,
				v('c:gapWidth', plot.gapWidth),
				container === 'barChart'
					? v('c:overlap', plot.overlap ?? (stacked ? 100 : undefined))
					: undefined,
				...axIds,
			];
			break;
		case 'lineChart':
		case 'stockChart':
			children = [
				container === 'lineChart'
					? valEl(
							'c:grouping',
							plot.grouping === 'clustered' ? 'standard' : (plot.grouping ?? 'standard'),
						)
					: undefined,
				container === 'lineChart' ? vary : undefined,
				...series,
				labels,
				plot.hiLowLines ? el('c:hiLowLines') : undefined,
				plot.upDownBars
					? el('c:upDownBars', {}, [valEl('c:gapWidth', 150), el('c:upBars'), el('c:downBars')])
					: undefined,
				container === 'lineChart' ? v('c:marker', plot.marker) : undefined,
				...axIds,
			];
			break;
		case 'line3DChart':
		case 'areaChart':
		case 'area3DChart':
			children = [
				valEl(
					'c:grouping',
					plot.grouping === 'clustered' ? 'standard' : (plot.grouping ?? 'standard'),
				),
				vary,
				...series,
				labels,
				...axIds,
			];
			break;
		case 'pieChart':
		case 'doughnutChart':
			children = [
				vary,
				...series,
				labels,
				v(
					'c:firstSliceAng',
					plot.firstSliceAngle === undefined ? undefined : Math.round(plot.firstSliceAngle),
				),
				container === 'doughnutChart'
					? v(
							'c:holeSize',
							plot.holeSize === undefined
								? undefined
								: Math.min(90, Math.max(1, Math.round(plot.holeSize))),
						)
					: undefined,
			];
			break;
		case 'pie3DChart':
			children = [vary, ...series, labels];
			break;
		case 'ofPieChart':
			children = [
				valEl('c:ofPieType', 'pie'),
				vary,
				...series,
				labels,
				v('c:gapWidth', plot.gapWidth),
			];
			break;
		case 'scatterChart':
			children = [
				valEl('c:scatterStyle', plot.scatterStyle ?? 'lineMarker'),
				vary,
				...series,
				labels,
				...axIds,
			];
			break;
		case 'radarChart':
			children = [
				valEl('c:radarStyle', plot.radarStyle ?? 'marker'),
				vary,
				...series,
				labels,
				...axIds,
			];
			break;
		case 'bubbleChart':
			children = [
				vary,
				...series,
				labels,
				v(
					'c:bubbleScale',
					plot.bubbleScale === undefined ? undefined : Math.round(plot.bubbleScale),
				),
				v('c:sizeRepresents', plot.bubbleSizeRepresents),
				...axIds,
			];
			break;
		default:
			children = [...series, ...axIds];
	}
	const out = el(`c:${container}`, {}, children);
	return mergeRawChildren(out, raw, {
		order: container,
		owned: [
			'barDir',
			'grouping',
			'varyColors',
			'ser',
			'dLbls',
			'gapWidth',
			'overlap',
			'axId',
			'firstSliceAng',
			'holeSize',
			'scatterStyle',
			'radarStyle',
			'marker',
			'hiLowLines',
			'upDownBars',
			'bubbleScale',
			'sizeRepresents',
		],
		preferRaw: ['ofPieType', 'hiLowLines', 'upDownBars'],
	});
}

function writeStyle(style: number | undefined): XmlElement | undefined {
	if (style === undefined) {
		return undefined;
	}
	if (style > 100) {
		return el('mc:AlternateContent', { 'xmlns:mc': NS.mc }, [
			el(
				'mc:Choice',
				{
					Requires: 'c14',
					'xmlns:c14': 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
				},
				[valEl('c14:style', style)],
			),
			el('mc:Fallback', {}, [valEl('c:style', style - 100)]),
		]);
	}
	return valEl('c:style', style);
}

function defaultView3D(plots: PlannedPlot[]): XmlElement | undefined {
	const p3d = plots.find((p) => p.plot.is3D);
	if (!p3d) {
		return undefined;
	}
	const pie = p3d.container === 'pie3DChart';
	return el('c:view3D', {}, [
		valEl('c:rotX', pie ? 30 : 15),
		valEl('c:rotY', pie ? 0 : 20),
		valEl('c:rAngAx', !pie),
	]);
}

function parseRaw(xml: string | undefined): XmlElement | undefined {
	if (!xml) {
		return undefined;
	}
	try {
		return parseXmlRoot(xml);
	} catch {
		return undefined;
	}
}

/** Build the `c:chartSpace` element for a model (merging unknown markup from `raw`). */
export function buildChartSpace(chart: ChartModel, raw?: XmlElement): XmlElement {
	const rawChart = rawChild(raw, 'chart');
	const rawPlotArea = rawChild(rawChart, 'plotArea');
	const planned = planPlots(chart);
	const seen = new Map<string, number>();
	const plotEls = planned.plots.map((p) => {
		const n = seen.get(p.container) ?? 0;
		seen.set(p.container, n + 1);
		return writePlot(p, rawChildren(rawPlotArea, p.container)[n]);
	});
	const axisEls = planned.axes.map((a) =>
		writeAxis(a, rawChildByKey(rawPlotArea, `${a.kind}Ax`, 'axId', a.id)),
	);
	const plotArea = mergeRawChildren(
		el('c:plotArea', {}, [
			writeLayout(chart.plotAreaLayout),
			...plotEls,
			...axisEls,
			writeShapeProps(chart.plotFill, chart.plotLine),
		]),
		rawPlotArea,
		{
			order: 'plotArea',
			owned: ['layout', 'spPr', ...CHART_ORDER.plotArea.filter((n) => /Chart$|Ax$/.test(n))],
		},
	);
	const chartEl = mergeRawChildren(
		el('c:chart', {}, [
			writeTitle(chart.title, rawChild(rawChart, 'title')),
			chart.autoTitleDeleted !== undefined
				? valEl('c:autoTitleDeleted', chart.autoTitleDeleted)
				: undefined,
			defaultView3D(planned.plots),
			plotArea,
			writeLegend(chart.legend, rawChild(rawChart, 'legend')),
			chart.plotVisibleOnly !== undefined
				? valEl('c:plotVisOnly', chart.plotVisibleOnly)
				: undefined,
			chart.displayBlanksAs ? valEl('c:dispBlanksAs', chart.displayBlanksAs) : undefined,
		]),
		rawChart,
		{
			order: 'chart',
			owned: ['title', 'autoTitleDeleted', 'plotArea', 'legend', 'plotVisOnly', 'dispBlanksAs'],
			preferRaw: ['view3D'],
		},
	);
	const attrs: Record<string, string> = { 'xmlns:c': NS.c, 'xmlns:a': NS.a, 'xmlns:r': NS.r };
	for (const [k, val] of Object.entries(raw?.attrs ?? {})) {
		if (k.startsWith('xmlns:') && !(k in attrs)) {
			attrs[k] = val;
		}
	}
	const space = el('c:chartSpace', attrs, [
		valEl('c:date1904', false),
		valEl('c:lang', 'en-US'),
		valEl('c:roundedCorners', chart.roundedCorners ?? false),
		writeStyle(chart.style),
		chartEl,
		writeShapeProps(chart.chartFill, chart.chartLine),
		writeTextProps(chart.textStyle),
	]);
	return mergeRawChildren(space, raw, {
		order: 'chartSpace',
		owned: ['roundedCorners', 'AlternateContent', 'style', 'chart', 'spPr', 'txPr'],
		preferRaw: ['date1904', 'lang'],
	});
}

/** True when the original part is a 2014 chartex (`cx:`) chart, which the model cannot regenerate. */
function isChartEx(raw: XmlElement | undefined): boolean {
	if (!raw) {
		return false;
	}
	return Object.values(raw.attrs).some((v) => v.includes('/chartex')) && !rawChild(raw, 'chart');
}

/** Serialize a chart model to a chart part. Unedited charts return their original XML. */
export function writeChartXml(chart: ChartModel): string {
	if (chart.rawXml && !chart.dirty) {
		return chart.rawXml;
	}
	const raw = parseRaw(chart.rawXml);
	if (raw && (localName(raw.name) !== 'chartSpace' || isChartEx(raw)) && chart.rawXml) {
		return chart.rawXml;
	}
	return serializeXml(
		buildChartSpace(chart, raw && raw.children.some(isElement) ? raw : undefined),
	);
}

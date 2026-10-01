/**
 * Chart layout: ChartModel + resolved data -> render tree. Allocates the chart
 * area like Excel (title on top, legend on its side, the plot area in the
 * remaining space) and delegates the plot area to the plot renderers.
 */
import type { ChartManualLayout, ChartModel } from '../../types/chart';
import type { DrawingParagraph } from '../../types/drawing';
import type { Theme } from '../../types/theme';
import type { ChartData } from '../data/types';
import { layoutCartesian } from './cartesian';
import {
	DEFAULT_GRIDLINE_COLOR,
	type LayoutContext,
	type ResolvedText,
	colorCss,
	createLayoutContext,
	fillPaint,
	resolveText,
	strokePaint,
} from './context';
import { chartTitleText, describeChart } from './describe';
import { drawLegend, legendEntries, planLegend } from './legend';
import { drawPiePlot } from './plot-pie';
import { drawRadarPlot } from './plot-radar';
import { blockHeight, textBlock } from './text-block';
import { type ChartTextMeasurer, wrapText } from './text';
import type {
	ChartBox,
	ChartGroupNode,
	ChartRenderNode,
	ChartRenderTree,
	ChartTextRun,
} from './types';

export interface ChartLayoutOptions {
	/** Text measurer (e.g. canvas based); defaults to built-in Calibri metrics. */
	measureText?: ChartTextMeasurer;
}

const GAP = 6;
const PIE_FAMILY = new Set(['pie', 'doughnut', 'ofPie']);

function manualBox(
	layout: ChartManualLayout | undefined,
	width: number,
	height: number,
	auto: ChartBox,
): ChartBox | undefined {
	if (!layout || layout.x === undefined || layout.y === undefined) {
		return undefined;
	}
	const edgeX = layout.xMode !== 'factor';
	const edgeY = layout.yMode !== 'factor';
	const x = edgeX ? layout.x * width : auto.x + layout.x * width;
	const y = edgeY ? layout.y * height : auto.y + layout.y * height;
	const w =
		layout.w !== undefined
			? edgeX
				? layout.w * width
				: auto.width + layout.w * width
			: auto.width;
	const h =
		layout.h !== undefined
			? edgeY
				? layout.h * height
				: auto.height + layout.h * height
			: auto.height;
	return { x, y, width: Math.max(10, w), height: Math.max(10, h) };
}

function richRuns(
	ctx: LayoutContext,
	base: ResolvedText,
	paragraphs: DrawingParagraph[],
): ChartTextRun[][] {
	return paragraphs.map((p) =>
		p.runs
			.filter((r) => r.text !== '\n')
			.map((r) => {
				const run: ChartTextRun = { text: r.text };
				const s = r.style;
				if (s?.bold !== undefined) {
					run.fontWeight = s.bold ? 'bold' : 'normal';
				}
				if (s?.italic !== undefined) {
					run.fontStyle = s.italic ? 'italic' : 'normal';
				}
				if (s?.underline) {
					run.decoration = 'underline';
				}
				if (s?.color) {
					run.fill = colorCss(ctx, s.color, base.color);
				}
				if (s?.size !== undefined) {
					run.fontSize = (s.size * 96) / 72;
				}
				return run;
			}),
	);
}

/** Draw the chart title at the top of `area`; returns the space it takes. */
function drawTitle(
	ctx: LayoutContext,
	area: ChartBox,
	width: number,
	height: number,
	out: ChartRenderNode[],
): number {
	const title = ctx.chart.title;
	const text = chartTitleText(ctx.chart, ctx.data);
	if (!title || !text) {
		return 0;
	}
	const style = resolveText(ctx, { sizePt: 14, relative: 1.4 }, title.style);
	const paragraphs = title.paragraphs?.filter((p) => p.runs.length) ?? [];
	const hasRichFormatting = paragraphs.some((p) => p.runs.some((r) => r.style));
	let lines = wrapText(text, area.width * 0.9, style.font, ctx.measure, 3);
	const extra: { runs?: ChartTextRun[][] } = {};
	if (hasRichFormatting && paragraphs.length === lines.length) {
		extra.runs = richRuns(ctx, style, paragraphs);
		lines = paragraphs.map((p) => p.runs.map((r) => r.text).join(''));
	}
	const h = blockHeight(style, lines.length);
	const manual = manualBox(title.layout, width, height, {
		x: area.x,
		y: area.y,
		width: area.width,
		height: h,
	});
	const meta = { role: 'title' as const };
	if (manual) {
		out.push(textBlock(style, lines, manual.x, manual.y, 'start', meta, extra));
	} else {
		out.push(textBlock(style, lines, area.x + area.width / 2, area.y, 'middle', meta, extra));
	}
	return title.overlay || manual ? 0 : h + GAP;
}

/** Place the legend inside `area` and return the area left for the plot. */
function drawLegendInto(
	ctx: LayoutContext,
	area: ChartBox,
	width: number,
	height: number,
	out: ChartRenderNode[],
): ChartBox {
	const legend = ctx.chart.legend;
	if (!legend) {
		return area;
	}
	const pos = legend.position;
	const vertical = pos === 'r' || pos === 'l' || pos === 'tr';
	const entries = legendEntries(ctx, vertical);
	if (!entries.length) {
		return area;
	}
	const plan = planLegend(
		ctx,
		entries,
		vertical,
		vertical ? area.width * 0.45 : area.width,
		vertical ? area.height : area.height * 0.4,
	);
	let x: number;
	let y: number;
	const next = { ...area };
	switch (pos) {
		case 'l':
			x = area.x;
			y = area.y + (area.height - plan.height) / 2;
			next.x += plan.width + GAP;
			next.width -= plan.width + GAP;
			break;
		case 't':
			x = area.x + (area.width - plan.width) / 2;
			y = area.y;
			next.y += plan.height + GAP;
			next.height -= plan.height + GAP;
			break;
		case 'b':
			x = area.x + (area.width - plan.width) / 2;
			y = area.y + area.height - plan.height;
			next.height -= plan.height + GAP;
			break;
		case 'tr':
			x = area.x + area.width - plan.width;
			y = area.y;
			next.width -= plan.width + GAP;
			break;
		default:
			x = area.x + area.width - plan.width;
			y = area.y + (area.height - plan.height) / 2;
			next.width -= plan.width + GAP;
	}
	const manual = manualBox(legend.layout, width, height, {
		x,
		y,
		width: plan.width,
		height: plan.height,
	});
	out.push(...drawLegend(ctx, plan, manual?.x ?? x, manual?.y ?? y));
	if (legend.overlay || manual) {
		return area;
	}
	next.width = Math.max(10, next.width);
	next.height = Math.max(10, next.height);
	return next;
}

/** Compute the render tree of a chart at `width` x `height` pixels. */
export function layoutChart(
	chart: ChartModel,
	data: ChartData,
	theme: Theme,
	width: number,
	height: number,
	options: ChartLayoutOptions = {},
): ChartRenderTree {
	const ctx = createLayoutContext(chart, data, theme, options.measureText);
	const W = Math.max(1, width);
	const H = Math.max(1, height);
	const root: ChartGroupNode = { type: 'group', children: [] };
	const frame = strokePaint(ctx, chart.chartLine, { color: DEFAULT_GRIDLINE_COLOR, width: 1 });
	const inset = (frame.strokeWidth ?? 0) / 2;
	root.children.push({
		type: 'rect',
		x: inset,
		y: inset,
		width: Math.max(0, W - 2 * inset),
		height: Math.max(0, H - 2 * inset),
		rx: chart.roundedCorners ? Math.min(10, Math.min(W, H) * 0.04) : undefined,
		...fillPaint(ctx, chart.chartFill, '#FFFFFF'),
		...frame,
		role: 'chart-area',
	});
	const pad = Math.max(4, Math.min(10, Math.min(W, H) * 0.025));
	let area: ChartBox = {
		x: pad,
		y: pad,
		width: Math.max(10, W - 2 * pad),
		height: Math.max(10, H - 2 * pad),
	};
	const overlays: ChartRenderNode[] = [];
	const used = drawTitle(ctx, area, W, H, overlays);
	area = { ...area, y: area.y + used, height: Math.max(10, area.height - used) };
	area = drawLegendInto(ctx, area, W, H, overlays);
	const manual = manualBox(chart.plotAreaLayout, W, H, area);
	const inner = chart.plotAreaLayout?.layoutTarget === 'inner' ? manual : undefined;
	const plotBox = inner ? area : (manual ?? area);
	const labels: ChartRenderNode[] = [];
	const plotNodes: ChartRenderNode[] = [];
	const chartBox: ChartBox = {
		x: pad,
		y: pad,
		width: Math.max(10, W - 2 * pad),
		height: Math.max(10, H - 2 * pad),
	};
	const cartesian = layoutCartesian(ctx, inner ? chartBox : plotBox, inner);
	if (cartesian) {
		plotNodes.push(...cartesian.nodes);
	}
	chart.plots.forEach((plot, index) => {
		const box = cartesian?.inner ?? inner ?? plotBox;
		if (PIE_FAMILY.has(plot.type)) {
			plotNodes.push(...drawPiePlot(ctx, box, plot, index, labels));
		} else if (plot.type === 'radar') {
			plotNodes.push(...drawRadarPlot(ctx, box, plot, index, labels));
		}
	});
	if (!chart.plots.length) {
		const style = resolveText(ctx, { sizePt: 10, color: '#7F7F7F' });
		plotNodes.push(
			textBlock(
				style,
				['This chart type is not supported'],
				plotBox.x + plotBox.width / 2,
				plotBox.y + plotBox.height / 2 - style.lineHeight / 2,
				'middle',
				{
					role: 'placeholder',
				},
			),
		);
	}
	root.children.push(...plotNodes, ...labels, ...overlays);
	return { width: W, height: H, root, description: describeChart(chart, data) };
}

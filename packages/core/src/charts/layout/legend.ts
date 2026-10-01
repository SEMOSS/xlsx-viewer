/** Legend entries (series, or categories for vary-colors charts), sizing and drawing. */
import type { ChartPlot } from '../../types/chart';
import {
	type LayoutContext,
	type ResolvedText,
	fillPaint,
	resolveText,
	strokePaint,
} from './context';
import { markerNode } from './geometry';
import {
	type ResolvedMarker,
	linePaint,
	orderedSeries,
	resolveMarker,
	seriesAutoColor,
	shapePaint,
	variesByPoint,
} from './series-style';
import { isStacked } from './stacking';
import { textBlock } from './text-block';
import { ellipsize } from './text';
import { trendlineLabel } from './trendline';
import type { ChartNodeMeta, ChartPaint, ChartRenderNode } from './types';

export interface LegendEntry {
	text: string;
	key: 'rect' | 'line';
	paint: ChartPaint;
	marker?: ResolvedMarker;
	meta: ChartNodeMeta;
}

const PAD = 4;
const KEY_GAP = 4;
const ENTRY_GAP = 12;

const FILLED = new Set(['bar', 'area', 'surface', 'bubble', 'pie', 'doughnut', 'ofPie']);

function reversesInVerticalLegend(plot: ChartPlot): boolean {
	return (
		(plot.type === 'bar' && (isStacked(plot) || plot.barDirection === 'bar')) ||
		(plot.type === 'area' && isStacked(plot))
	);
}

/** Legend entries in Excel's order, without deleted entries. */
export function legendEntries(ctx: LayoutContext, vertical: boolean): LegendEntry[] {
	const { chart, data } = ctx;
	const deleted = new Set(chart.legend?.deletedEntries ?? []);
	const entries: LegendEntry[] = [];
	const single = chart.plots.length === 1 ? chart.plots[0] : undefined;
	if (single && variesByPoint(single)) {
		const visible = orderedSeries(single).find(
			({ index }) => !data.plots[0]?.series[index]?.hidden,
		);
		const sd = visible ? data.plots[0].series[visible.index] : undefined;
		if (visible && sd) {
			const count = Math.max(sd.categories.length, sd.values.length);
			for (let i = 0; i < count; i++) {
				if (deleted.has(i)) {
					continue;
				}
				entries.push({
					text: sd.categories[i] ?? String(i + 1),
					key: 'rect',
					paint: shapePaint(ctx, single, visible.series, i, count),
					meta: { role: 'legend-entry', plot: 0, series: visible.index, point: i },
				});
			}
		}
		return entries;
	}
	const trends: LegendEntry[] = [];
	chart.plots.forEach((plot, p) => {
		const group: LegendEntry[] = [];
		for (const { series, index } of orderedSeries(plot)) {
			const sd = data.plots[p]?.series[index];
			if (!sd || sd.hidden || deleted.has(series.idx)) {
				continue;
			}
			const filled =
				FILLED.has(plot.type) || (plot.type === 'radar' && plot.radarStyle === 'filled');
			const meta: ChartNodeMeta = { role: 'legend-entry', plot: p, series: index };
			if (filled) {
				group.push({ text: sd.name, key: 'rect', paint: shapePaint(ctx, plot, series), meta });
			} else {
				const entry: LegendEntry = {
					text: sd.name,
					key: 'line',
					paint: linePaint(
						ctx,
						plot,
						series,
						plot.type !== 'stock' && !(plot.type === 'scatter' && plot.scatterStyle === 'marker'),
					),
					meta,
				};
				const marker = resolveMarker(ctx, plot, series);
				if (marker) {
					entry.marker = marker;
				}
				group.push(entry);
			}
			for (const trend of series.trendlines ?? []) {
				trends.push({
					text: trendlineLabel(trend, sd.name),
					key: 'line',
					paint: strokePaint(ctx, trend.line, {
						color: seriesAutoColor(ctx, plot, series),
						width: 2,
					}),
					meta: { role: 'legend-entry', plot: p, series: index },
				});
			}
		}
		entries.push(...(vertical && reversesInVerticalLegend(plot) ? group.reverse() : group));
	});
	return [...entries, ...trends];
}

export interface LegendPlan {
	entries: LegendEntry[];
	style: ResolvedText;
	width: number;
	height: number;
	/** Entry rows (horizontal legends) or one entry per row (vertical legends). */
	rows: LegendEntry[][];
	/** Fixed column width for grid layouts; 0 for natural widths. */
	columnWidth: number;
	texts: Map<LegendEntry, string>;
}

function keyWidth(entry: LegendEntry, style: ResolvedText): number {
	return entry.key === 'line' ? style.sizePx * 1.9 : Math.max(5, Math.round(style.sizePx * 0.6));
}

/** Size the legend within the available space. */
export function planLegend(
	ctx: LayoutContext,
	entries: LegendEntry[],
	vertical: boolean,
	maxWidth: number,
	maxHeight: number,
): LegendPlan {
	const style = resolveText(ctx, { sizePt: 10 }, ctx.chart.legend?.textStyle);
	const texts = new Map<LegendEntry, string>();
	const widths = new Map<LegendEntry, number>();
	const textCap = Math.max(20, (vertical ? maxWidth : maxWidth * 0.9) - 2 * PAD);
	for (const e of entries) {
		const kw = keyWidth(e, style) + KEY_GAP;
		const text = ellipsize(e.text, Math.max(10, textCap - kw), style.font, ctx.measure);
		texts.set(e, text);
		widths.set(e, kw + ctx.measure(text, style.font));
	}
	const rowH = Math.max(style.lineHeight, style.sizePx * 0.6);
	const widest = Math.max(0, ...entries.map((e) => widths.get(e) ?? 0));
	if (vertical) {
		const fit = Math.max(1, Math.floor((maxHeight - 2 * PAD) / rowH));
		const shown = entries.slice(0, fit);
		return {
			entries: shown,
			style,
			width: Math.min(maxWidth, widest + 2 * PAD),
			height: shown.length * rowH + 2 * PAD,
			rows: shown.map((e) => [e]),
			columnWidth: 0,
			texts,
		};
	}
	const natural =
		entries.reduce((s, e) => s + (widths.get(e) ?? 0), 0) +
		ENTRY_GAP * Math.max(0, entries.length - 1);
	if (natural + 2 * PAD <= maxWidth) {
		return {
			entries,
			style,
			width: natural + 2 * PAD,
			height: rowH + 2 * PAD,
			rows: [entries],
			columnWidth: 0,
			texts,
		};
	}
	const colW = widest + ENTRY_GAP;
	const cols = Math.max(1, Math.floor((maxWidth - 2 * PAD + ENTRY_GAP) / colW));
	const maxRows = Math.max(1, Math.floor((maxHeight - 2 * PAD) / rowH));
	const rows: LegendEntry[][] = [];
	for (let i = 0; i < entries.length && rows.length < maxRows; i += cols) {
		rows.push(entries.slice(i, i + cols));
	}
	const used = Math.min(cols, entries.length);
	return {
		entries: rows.flat(),
		style,
		width: Math.min(maxWidth, used * colW - ENTRY_GAP + 2 * PAD),
		height: rows.length * rowH + 2 * PAD,
		rows,
		columnWidth: colW,
		texts,
	};
}

/** Draw a planned legend with its top-left corner at (x, y). */
export function drawLegend(
	ctx: LayoutContext,
	plan: LegendPlan,
	x: number,
	y: number,
): ChartRenderNode[] {
	const legend = ctx.chart.legend;
	const out: ChartRenderNode[] = [];
	const boxPaint = {
		...fillPaint(ctx, legend?.fill, null),
		...strokePaint(ctx, legend?.line, null),
	};
	if (boxPaint.fill || boxPaint.gradient || boxPaint.stroke) {
		out.push({
			type: 'rect',
			x,
			y,
			width: plan.width,
			height: plan.height,
			...boxPaint,
			role: 'legend',
		});
	}
	const style = plan.style;
	const rowH = Math.max(style.lineHeight, style.sizePx * 0.6);
	plan.rows.forEach((row, r) => {
		const rowWidth =
			plan.columnWidth > 0
				? row.length * plan.columnWidth - ENTRY_GAP
				: row.reduce(
						(s, e) =>
							s + keyWidth(e, style) + KEY_GAP + ctx.measure(plan.texts.get(e) ?? '', style.font),
						0,
					) +
					ENTRY_GAP * (row.length - 1);
		let cx =
			plan.rows.length === 1 && plan.columnWidth === 0 ? x + (plan.width - rowWidth) / 2 : x + PAD;
		const cy = y + PAD + r * rowH + rowH / 2;
		for (const entry of row) {
			const kw = keyWidth(entry, style);
			out.push({
				type: 'group',
				role: 'legend-entry',
				...entry.meta,
				children: legendKey(entry, cx, cy, kw, style),
			});
			const text = plan.texts.get(entry) ?? entry.text;
			out.push(
				textBlock(style, [text], cx + kw + KEY_GAP, cy - style.lineHeight / 2, 'start', {
					...entry.meta,
					role: 'legend-entry',
				}),
			);
			const advance = kw + KEY_GAP + ctx.measure(text, style.font) + ENTRY_GAP;
			cx += plan.columnWidth > 0 ? plan.columnWidth : advance;
		}
	});
	return out;
}

function legendKey(
	entry: LegendEntry,
	x: number,
	cy: number,
	width: number,
	style: ResolvedText,
): ChartRenderNode[] {
	if (entry.key === 'rect') {
		const size = width;
		return [
			{
				type: 'rect',
				x,
				y: cy - size / 2,
				width: size,
				height: size,
				...entry.paint,
				role: 'legend-key',
			},
		];
	}
	const out: ChartRenderNode[] = [];
	if (entry.paint.stroke) {
		out.push({
			type: 'line',
			x1: x,
			y1: cy,
			x2: x + width,
			y2: cy,
			...entry.paint,
			strokeWidth: Math.min(entry.paint.strokeWidth ?? 1, 4),
			role: 'legend-key',
		});
	}
	if (entry.marker) {
		const size = Math.min(entry.marker.size, style.lineHeight * 0.8);
		const node = markerNode(entry.marker.symbol, x + width / 2, cy, size, entry.marker.paint);
		if (node) {
			out.push({ ...node, role: 'legend-key' } as ChartRenderNode);
		}
	}
	return out;
}

/** Line, area, stock and (approximated) surface plots on a category axis. */
import type { ChartPlot, ChartSeries } from '../../types/chart';
import type { ChartSeriesData } from '../data/types';
import {
	type AxisGroup,
	categoryFraction,
	catPixel,
	crossValue,
	slotFraction,
	valPixel,
} from './axis-model';
import { type LayoutContext, PX_PER_PT, fillPaint, strokePaint } from './context';
import { pointLabel } from './data-labels';
import { type Point, linePath, markerNode, smoothPath } from './geometry';
import {
	linePaint,
	orderedSeries,
	resolveMarker,
	seriesAutoColor,
	shapePaint,
} from './series-style';
import { isStacked, stackValues } from './stacking';
import { fitTrendline, movingAverage } from './trendline';
import type { ChartBox, ChartPaint, ChartRenderNode } from './types';

/** Split points into drawable runs: `null` breaks the line, `NaN` points are skipped. */
export function lineRuns(points: ReadonlyArray<{ x: number; y: number | null }>): Point[][] {
	const runs: Point[][] = [];
	let cur: Point[] = [];
	for (const p of points) {
		if (p.y === null) {
			if (cur.length) {
				runs.push(cur);
			}
			cur = [];
		} else if (Number.isFinite(p.y)) {
			cur.push([p.x, p.y]);
		}
	}
	if (cur.length) {
		runs.push(cur);
	}
	return runs;
}

/** Path node(s) for runs of points. */
export function runsPath(runs: Point[][], smooth: boolean): string {
	return runs
		.filter((r) => r.length > 1)
		.map((r) => (smooth ? smoothPath(r) : linePath(r)))
		.join(' ');
}

/** Trendlines of a series, drawn across the X extent of its data. */
export function drawTrendlines(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	xs: number[],
	ys: number[],
	toPixel: (x: number, y: number) => Point,
	meta: { plot: number; series: number },
): ChartRenderNode[] {
	const out: ChartRenderNode[] = [];
	const color = seriesAutoColor(ctx, plot, series);
	for (const trend of series.trendlines ?? []) {
		const paint = strokePaint(ctx, trend.line, { color, width: 1.5 * PX_PER_PT });
		if (!trend.line?.dash && paint.strokeWidth) {
			paint.dash = [paint.strokeWidth, paint.strokeWidth];
		}
		let pts: Point[];
		if (trend.type === 'movingAvg') {
			pts = movingAverage(xs, ys, trend.period ?? 2).map(([x, y]) => toPixel(x, y));
		} else {
			const fn = fitTrendline(trend, xs, ys);
			const finite = xs.filter((x) => Number.isFinite(x));
			if (!fn || finite.length < 2) {
				continue;
			}
			const lo = Math.min(...finite);
			const hi = Math.max(...finite);
			const steps = trend.type === 'linear' ? 1 : 60;
			pts = [];
			for (let k = 0; k <= steps; k++) {
				const x = lo + ((hi - lo) * k) / steps;
				const y = fn(x);
				if (Number.isFinite(y)) {
					pts.push(toPixel(x, y));
				}
			}
		}
		if (pts.length > 1) {
			out.push({ type: 'path', d: linePath(pts), ...paint, role: 'trendline', ...meta });
		}
	}
	return out;
}

function markersFor(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	pts: ReadonlyArray<{ x: number; y: number | null; i: number }>,
	meta: { plot: number; series: number },
): ChartRenderNode[] {
	const out: ChartRenderNode[] = [];
	for (const p of pts) {
		if (p.y === null || !Number.isFinite(p.y)) {
			continue;
		}
		const m = resolveMarker(ctx, plot, series, p.i);
		const node = m ? markerNode(m.symbol, p.x, p.y, m.size, m.paint) : undefined;
		if (node) {
			out.push({ ...node, ...meta, point: p.i } as ChartRenderNode);
		}
	}
	return out;
}

function info(data: ChartSeriesData, i: number, value: number | null, percent?: number) {
	return {
		value,
		category: data.categories[i] ?? String(i + 1),
		seriesName: data.name,
		percent,
		valueFormat: data.valueFormat,
	};
}

/** Line, stock-series and surface plots: one polyline per series plus markers. */
export function drawLinePlot(
	ctx: LayoutContext,
	group: AxisGroup,
	inner: ChartBox,
	plot: ChartPlot,
	plotIndex: number,
	labels: ChartRenderNode[],
): ChartRenderNode[] {
	const plotData = ctx.data.plots[plotIndex];
	if (!plotData) {
		return [];
	}
	const stacked = isStacked(plot);
	const stacks = stackValues(plot, plotData);
	const out: ChartRenderNode[] = [];
	if (plot.type === 'stock') {
		out.push(...drawStockDecorations(ctx, group, inner, plot, plotIndex));
	}
	for (const { series, index } of orderedSeries(plot)) {
		const data = plotData.series[index];
		if (!data || data.hidden) {
			continue;
		}
		const meta = { plot: plotIndex, series: index };
		const pts = (stacks[index] ?? []).map((sv, i) => {
			const x = catPixel(group, inner, categoryFraction(group.cat, i));
			const v = sv.raw === null ? null : !Number.isFinite(sv.raw) ? NaN : stacked ? sv.to : sv.raw;
			return {
				x,
				y: v === null ? null : Number.isNaN(v) ? NaN : valPixel(group, inner, v),
				i,
				share: sv.share,
				raw: sv.raw,
			};
		});
		const nodes: ChartRenderNode[] = [];
		const paint = linePaint(ctx, plot, series, plot.type !== 'stock');
		const d = runsPath(lineRuns(pts), Boolean(series.smooth));
		if (paint.stroke && d) {
			nodes.push({ type: 'path', d, ...paint, role: 'line', ...meta });
		}
		nodes.push(...markersFor(ctx, plot, series, pts, meta));
		const xs = pts.map((p) => p.i + 1);
		const ys = pts.map((p) => (p.raw === null ? NaN : (p.raw as number)));
		nodes.push(
			...drawTrendlines(
				ctx,
				plot,
				series,
				xs,
				ys,
				(x, y) => [
					catPixel(group, inner, categoryFraction(group.cat, x - 1)),
					valPixel(group, inner, y),
				],
				meta,
			),
		);
		out.push({ type: 'group', children: nodes, role: 'series', ...meta });
		for (const p of pts) {
			if (p.y !== null && Number.isFinite(p.y)) {
				const m = resolveMarker(ctx, plot, series, p.i);
				labels.push(
					...pointLabel(
						ctx,
						plot,
						series,
						p.i,
						p.x,
						p.y,
						m ? m.size / 2 : 0,
						info(data, p.i, p.raw, p.share),
						'r',
						{
							...meta,
							point: p.i,
						},
					),
				);
			}
		}
	}
	return out;
}

/** Area plots: filled polygons down to the category axis (or the previous stack). */
export function drawAreaPlot(
	ctx: LayoutContext,
	group: AxisGroup,
	inner: ChartBox,
	plot: ChartPlot,
	plotIndex: number,
	labels: ChartRenderNode[],
): ChartRenderNode[] {
	const plotData = ctx.data.plots[plotIndex];
	if (!plotData) {
		return [];
	}
	const stacked = isStacked(plot);
	const base = stacked ? 0 : crossValue(group);
	const stacks = stackValues(plot, plotData);
	const out: ChartRenderNode[] = [];
	for (const { series, index } of orderedSeries(plot)) {
		const data = plotData.series[index];
		const points = stacks[index] ?? [];
		if (!data || data.hidden || points.length === 0) {
			continue;
		}
		const meta = { plot: plotIndex, series: index };
		const xs = points.map((_, i) => catPixel(group, inner, categoryFraction(group.cat, i)));
		const value = (v: number | null) => (v === null || !Number.isFinite(v) ? 0 : v);
		const top: Point[] = points.map((sv, i) => [
			xs[i],
			valPixel(group, inner, stacked ? sv.to : value(sv.raw)),
		]);
		const bottom: Point[] = points
			.map((sv, i): Point => [xs[i], valPixel(group, inner, stacked ? sv.from : base)])
			.reverse();
		const d = `${linePath([...top, ...bottom])} Z`;
		const paint: ChartPaint = shapePaint(ctx, plot, series);
		out.push({
			type: 'group',
			role: 'series',
			...meta,
			children: [{ type: 'path', d, ...paint, role: 'area', ...meta }],
		});
		points.forEach((sv, i) => {
			if (sv.raw === null || !Number.isFinite(sv.raw)) {
				return;
			}
			const lo = valPixel(group, inner, stacked ? sv.from : base);
			const hi = valPixel(group, inner, stacked ? sv.to : sv.raw);
			labels.push(
				...pointLabel(
					ctx,
					plot,
					series,
					i,
					xs[i],
					(lo + hi) / 2,
					0,
					info(data, i, sv.raw, sv.share),
					'ctr',
					{
						...meta,
						point: i,
					},
				),
			);
		});
	}
	return out;
}

/** High-low lines and up/down bars of a stock plot. */
function drawStockDecorations(
	ctx: LayoutContext,
	group: AxisGroup,
	inner: ChartBox,
	plot: ChartPlot,
	plotIndex: number,
): ChartRenderNode[] {
	const plotData = ctx.data.plots[plotIndex];
	const series = plotData.series.filter((s) => !s.hidden);
	const out: ChartRenderNode[] = [];
	const count = group.cat.count;
	const finite = (v: number | null | undefined): v is number =>
		v !== null && v !== undefined && Number.isFinite(v);
	if (plot.hiLowLines) {
		const paint = strokePaint(ctx, undefined, { color: '#404040', width: 1 });
		for (let i = 0; i < count; i++) {
			const vals = series.map((s) => s.values[i]).filter(finite);
			if (vals.length < 2) {
				continue;
			}
			const x = catPixel(group, inner, categoryFraction(group.cat, i));
			out.push({
				type: 'line',
				x1: x,
				y1: valPixel(group, inner, Math.min(...vals)),
				x2: x,
				y2: valPixel(group, inner, Math.max(...vals)),
				...paint,
				role: 'hi-low-line',
				plot: plotIndex,
				point: i,
			});
		}
	}
	if (plot.upDownBars && series.length >= 2) {
		const open = series[0];
		const close = series[series.length - 1];
		const width = slotFraction(group.cat) / 2.5;
		for (let i = 0; i < count; i++) {
			const o = open.values[i];
			const c = close.values[i];
			if (!finite(o) || !finite(c)) {
				continue;
			}
			const f = categoryFraction(group.cat, i);
			const x0 = catPixel(group, inner, f - width / 2);
			const x1 = catPixel(group, inner, f + width / 2);
			const y0 = valPixel(group, inner, o);
			const y1 = valPixel(group, inner, c);
			const up = c >= o;
			out.push({
				type: 'rect',
				x: Math.min(x0, x1),
				y: Math.min(y0, y1),
				width: Math.abs(x1 - x0),
				height: Math.max(1, Math.abs(y1 - y0)),
				...fillPaint(ctx, undefined, up ? '#FFFFFF' : '#595959'),
				stroke: '#595959',
				strokeWidth: 1,
				role: 'up-down-bar',
				plot: plotIndex,
				point: i,
			});
		}
	}
	return out;
}

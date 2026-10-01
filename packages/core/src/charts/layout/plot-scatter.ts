/** XY scatter and bubble plots. */
import type { ChartPlot } from '../../types/chart';
import { type AxisGroup, valPixel, xPixel } from './axis-model';
import type { LayoutContext } from './context';
import { pointLabel } from './data-labels';
import { markerNode } from './geometry';
import { drawTrendlines, lineRuns, runsPath } from './plot-line';
import { linePaint, orderedSeries, resolveMarker, shapePaint } from './series-style';
import { xValueAt } from './stacking';
import type { ChartBox, ChartRenderNode } from './types';

/** Whether a scatter plot draws connecting lines for series without explicit line formatting. */
function defaultLines(plot: ChartPlot): boolean {
	return plot.scatterStyle !== 'marker' && plot.scatterStyle !== 'none';
}

export function drawScatterPlot(
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
	const out: ChartRenderNode[] = [];
	for (const { series, index } of orderedSeries(plot)) {
		const data = plotData.series[index];
		if (!data || data.hidden) {
			continue;
		}
		const meta = { plot: plotIndex, series: index };
		const pts = data.values.map((v, i) => {
			const xv = xValueAt(data, i);
			const valid = v !== null && Number.isFinite(v) && Number.isFinite(xv);
			return {
				i,
				x: Number.isFinite(xv) ? xPixel(group, inner, xv) : NaN,
				y: v === null ? null : valid ? valPixel(group, inner, v) : NaN,
				xv,
				v,
			};
		});
		const nodes: ChartRenderNode[] = [];
		const paint = linePaint(ctx, plot, series, defaultLines(plot));
		const smooth =
			series.smooth ?? (plot.scatterStyle === 'smooth' || plot.scatterStyle === 'smoothMarker');
		const d = runsPath(
			lineRuns(pts.map((p) => ({ x: p.x, y: Number.isFinite(p.x) ? p.y : NaN }))),
			smooth,
		);
		if (paint.stroke && d) {
			nodes.push({ type: 'path', d, ...paint, role: 'line', ...meta });
		}
		for (const p of pts) {
			if (p.y === null || !Number.isFinite(p.y) || !Number.isFinite(p.x)) {
				continue;
			}
			const m = resolveMarker(ctx, plot, series, p.i);
			const node = m ? markerNode(m.symbol, p.x, p.y, m.size, m.paint) : undefined;
			if (node) {
				nodes.push({ ...node, ...meta, point: p.i } as ChartRenderNode);
			}
			labels.push(
				...pointLabel(
					ctx,
					plot,
					series,
					p.i,
					p.x,
					p.y,
					m ? m.size / 2 : 0,
					{
						value: p.v,
						category: data.categories[p.i] ?? String(p.xv),
						seriesName: data.name,
						valueFormat: data.valueFormat,
					},
					'r',
					{ ...meta, point: p.i },
				),
			);
		}
		nodes.push(
			...drawTrendlines(
				ctx,
				plot,
				series,
				pts.map((p) => p.xv),
				pts.map((p) => (p.v === null ? NaN : p.v)),
				(x, y) => [xPixel(group, inner, x), valPixel(group, inner, y)],
				meta,
			),
		);
		out.push({ type: 'group', children: nodes, role: 'series', ...meta });
	}
	return out;
}

export function drawBubblePlot(
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
	let maxSize = 0;
	for (const s of plotData.series) {
		for (const b of s.bubbleSizes ?? []) {
			if (b !== null && Number.isFinite(b)) {
				maxSize = Math.max(maxSize, Math.abs(b));
			}
		}
	}
	// The largest bubble spans a quarter of the plot's smaller side at 100% scale.
	const maxR =
		(Math.max(0, Math.min(300, plot.bubbleScale ?? 100)) / 100) *
		Math.min(inner.width, inner.height) *
		0.125;
	const out: ChartRenderNode[] = [];
	for (const { series, index } of orderedSeries(plot)) {
		const data = plotData.series[index];
		if (!data || data.hidden) {
			continue;
		}
		const meta = { plot: plotIndex, series: index };
		const nodes: ChartRenderNode[] = [];
		data.values.forEach((v, i) => {
			const xv = xValueAt(data, i);
			const size = data.bubbleSizes?.[i] ?? null;
			if (
				v === null ||
				!Number.isFinite(v) ||
				!Number.isFinite(xv) ||
				size === null ||
				!(size > 0) ||
				maxSize <= 0
			) {
				return;
			}
			const ratio = size / maxSize;
			const r = plot.bubbleSizeRepresents === 'w' ? maxR * ratio : maxR * Math.sqrt(ratio);
			const cx = xPixel(group, inner, xv);
			const cy = valPixel(group, inner, v);
			nodes.push({
				type: 'circle',
				cx,
				cy,
				r: Math.max(1, r),
				...shapePaint(ctx, plot, series, i, data.values.length),
				role: 'bubble',
				...meta,
				point: i,
			});
			labels.push(
				...pointLabel(
					ctx,
					plot,
					series,
					i,
					cx,
					cy,
					r,
					{
						value: v,
						category: data.categories[i] ?? String(xv),
						seriesName: data.name,
						bubbleSize: size,
						valueFormat: data.valueFormat,
					},
					'ctr',
					{ ...meta, point: i },
				),
			);
		});
		out.push({ type: 'group', children: nodes, role: 'series', ...meta });
	}
	return out;
}

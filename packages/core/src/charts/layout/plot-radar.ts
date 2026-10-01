/** Radar charts (standard, with markers, filled) on a polar grid. */
import type { ChartAxis, ChartPlot } from '../../types/chart';
import { axisLabelStyle } from './axis-labels';
import {
	DEFAULT_AXIS_COLOR,
	DEFAULT_GRIDLINE_COLOR,
	type LayoutContext,
	formatNumber,
	strokePaint,
} from './context';
import { pointLabel } from './data-labels';
import { type Point, linePath, markerNode, polar } from './geometry';
import { computeValueScale, maxIntervalsFor } from './scale';
import { linePaint, orderedSeries, resolveMarker, shapePaint } from './series-style';
import { plotValueRange, pointCount } from './stacking';
import { textBlock } from './text-block';
import type { ChartBox, ChartRenderNode } from './types';

const GAP = 5;

function radarAxes(ctx: LayoutContext, plot: ChartPlot): { cat: ChartAxis; val: ChartAxis } {
	const byId = new Map(ctx.chart.axes.map((a) => [a.id, a] as const));
	const a0 = byId.get(plot.axisIds[0]);
	const a1 = byId.get(plot.axisIds[1]);
	const cat = [a0, a1].find((a) => a && a.kind !== 'val') ?? {
		id: -1,
		kind: 'cat',
		position: 'b',
		crossAxisId: -2,
	};
	const val = [a0, a1].find((a) => a && a.kind === 'val') ?? {
		id: -2,
		kind: 'val',
		position: 'l',
		crossAxisId: -1,
	};
	return { cat, val };
}

export function drawRadarPlot(
	ctx: LayoutContext,
	box: ChartBox,
	plot: ChartPlot,
	plotIndex: number,
	labels: ChartRenderNode[],
): ChartRenderNode[] {
	const plotData = ctx.data.plots[plotIndex];
	if (!plotData) {
		return [];
	}
	const { cat, val } = radarAxes(ctx, plot);
	const n = pointCount(plotData);
	if (n === 0) {
		return [];
	}
	const catStyle = axisLabelStyle(ctx, cat);
	const valStyle = axisLabelStyle(ctx, val);
	const first = plotData.series.find((s) => !s.hidden);
	const catLabels = Array.from({ length: n }, (_, i) => first?.categories[i] ?? String(i + 1));
	const showCat = !cat.deleted && cat.tickLabelPosition !== 'none';
	const widest = showCat ? Math.max(...catLabels.map((l) => ctx.measure(l, catStyle.font))) : 0;
	const cx = box.x + box.width / 2;
	const cy = box.y + box.height / 2;
	const room = showCat
		? Math.min(box.width / 2 - widest - GAP, box.height / 2 - catStyle.lineHeight - GAP)
		: Math.min(box.width, box.height) / 2;
	const radius = Math.max(Math.min(box.width, box.height) * 0.2, room);
	const range = plotValueRange(plot, plotData);
	const scale = computeValueScale(range?.min ?? 0, range?.max ?? 0, {
		min: val.min,
		max: val.max,
		majorUnit: val.majorUnit,
		maxIntervals: maxIntervalsFor(radius, valStyle.lineHeight * 1.25),
	});
	const angle = (i: number) => (i * 360) / n;
	const rOf = (v: number) =>
		((Math.max(scale.min, Math.min(scale.max, v)) - scale.min) / (scale.max - scale.min)) * radius;
	const out: ChartRenderNode[] = [];
	// Web: one polygon per major unit, plus spokes along every category.
	if (val.majorGridlines) {
		const paint = strokePaint(ctx, val.majorGridlines === true ? undefined : val.majorGridlines, {
			color: DEFAULT_GRIDLINE_COLOR,
			width: 1,
		});
		for (const t of scale.ticks) {
			if (rOf(t) < 0.5) {
				continue;
			}
			const pts = Array.from({ length: n }, (_, i) => polar(cx, cy, rOf(t), angle(i)));
			out.push({ type: 'path', d: `${linePath(pts)} Z`, ...paint, role: 'gridline' });
		}
	}
	if (!cat.deleted) {
		const paint = strokePaint(ctx, cat.line, { color: DEFAULT_GRIDLINE_COLOR, width: 1 });
		for (let i = 0; i < n && paint.stroke; i++) {
			const [x, y] = polar(cx, cy, radius, angle(i));
			out.push({ type: 'line', x1: cx, y1: cy, x2: x, y2: y, ...paint, role: 'axis-line' });
		}
	}
	if (!val.deleted) {
		const paint = strokePaint(ctx, val.line, { color: DEFAULT_AXIS_COLOR, width: 1 });
		if (paint.stroke) {
			out.push({
				type: 'line',
				x1: cx,
				y1: cy,
				x2: cx,
				y2: cy - radius,
				...paint,
				role: 'axis-line',
			});
		}
		if (val.tickLabelPosition !== 'none') {
			const format =
				val.numFmt && !val.numFmt.sourceLinked
					? val.numFmt.code
					: (first?.valueFormat ?? 'General');
			for (const t of scale.ticks) {
				const y = cy - rOf(t);
				out.push(
					textBlock(
						valStyle,
						[formatNumber(ctx, t, format)],
						cx - GAP,
						y - valStyle.lineHeight / 2,
						'end',
						{
							role: 'tick-label',
						},
					),
				);
			}
		}
	}
	if (showCat) {
		catLabels.forEach((text, i) => {
			const a = angle(i);
			const [x, y] = polar(cx, cy, radius + GAP, a);
			const anchor = Math.abs(x - cx) < 1 ? 'middle' : x > cx ? 'start' : 'end';
			const top =
				Math.abs(y - cy) < 1 ? y - catStyle.lineHeight / 2 : y < cy ? y - catStyle.lineHeight : y;
			out.push(textBlock(catStyle, [text], x, top, anchor, { role: 'tick-label' }));
		});
	}
	const filled = plot.radarStyle === 'filled';
	for (const { series, index } of orderedSeries(plot)) {
		const data = plotData.series[index];
		if (!data || data.hidden) {
			continue;
		}
		const meta = { plot: plotIndex, series: index };
		const pts: Array<{ p: Point; i: number; v: number }> = [];
		data.values.forEach((v, i) => {
			if (v !== null && Number.isFinite(v)) {
				pts.push({ p: polar(cx, cy, rOf(v), angle(i)), i, v });
			}
		});
		if (!pts.length) {
			continue;
		}
		const nodes: ChartRenderNode[] = [];
		const d = `${linePath(pts.map((q) => q.p))} Z`;
		if (filled) {
			nodes.push({ type: 'path', d, ...shapePaint(ctx, plot, series), role: 'area', ...meta });
		} else {
			const paint = linePaint(ctx, plot, series);
			if (paint.stroke) {
				nodes.push({ type: 'path', d, ...paint, role: 'line', ...meta });
			}
		}
		for (const q of pts) {
			const m = filled ? undefined : resolveMarker(ctx, plot, series, q.i);
			const node = m ? markerNode(m.symbol, q.p[0], q.p[1], m.size, m.paint) : undefined;
			if (node) {
				nodes.push({ ...node, ...meta, point: q.i } as ChartRenderNode);
			}
			labels.push(
				...pointLabel(
					ctx,
					plot,
					series,
					q.i,
					q.p[0],
					q.p[1],
					m ? m.size / 2 : 0,
					{
						value: q.v,
						category: data.categories[q.i] ?? '',
						seriesName: data.name,
						valueFormat: data.valueFormat,
					},
					'r',
					{ ...meta, point: q.i },
				),
			);
		}
		out.push({ type: 'group', children: nodes, role: 'series', ...meta });
	}
	return out;
}

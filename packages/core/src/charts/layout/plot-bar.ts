/** Clustered, stacked and 100% stacked bars and columns. */
import type { ChartDataLabels, ChartPlot } from '../../types/chart';
import {
	type AxisGroup,
	categoryFraction,
	catPixel,
	crossValue,
	slotFraction,
	valPixel,
} from './axis-model';
import type { LayoutContext } from './context';
import { effectiveLabels, labelText, labelTextStyle, placeLabel } from './data-labels';
import { findPoint, orderedSeries, shapePaint } from './series-style';
import { isStacked, stackValues } from './stacking';
import type { ChartBox, ChartRenderNode, ChartTextNode } from './types';

const LABEL_OFFSET = 3;

/** Bar width and per-series offset, as fractions of the category axis. */
export function barGeometry(
	plot: ChartPlot,
	slot: number,
	seriesCount: number,
): { width: number; step: number; lead: number } {
	const gap = Math.max(0, Math.min(500, plot.gapWidth ?? 150)) / 100;
	if (isStacked(plot) || seriesCount <= 1) {
		const width = slot / (1 + gap);
		return { width, step: 0, lead: (width * gap) / 2 };
	}
	const overlap = Math.max(-1, Math.min(1, (plot.overlap ?? 0) / 100));
	const n = seriesCount;
	const width = slot / (n - (n - 1) * overlap + gap);
	return { width, step: width * (1 - overlap), lead: (width * gap) / 2 };
}

/** Label anchor for a bar given the label position (default: outside end, or centre for stacks). */
function barLabelAnchor(
	position: ChartDataLabels['position'],
	stacked: boolean,
	horizontal: boolean,
	rect: ChartBox,
	positive: boolean,
): { x: number; y: number; anchor: ChartTextNode['anchor']; vAlign: 'top' | 'middle' | 'bottom' } {
	const pos = position ?? (stacked ? 'ctr' : 'outEnd');
	const { x, y, width: w, height: h } = rect;
	if (horizontal) {
		const endX = positive ? x + w : x;
		const baseX = positive ? x : x + w;
		const dir = positive ? 1 : -1;
		const cy = y + h / 2;
		switch (pos) {
			case 'inEnd':
				return {
					x: endX - dir * LABEL_OFFSET,
					y: cy,
					anchor: positive ? 'end' : 'start',
					vAlign: 'middle',
				};
			case 'inBase':
				return {
					x: baseX + dir * LABEL_OFFSET,
					y: cy,
					anchor: positive ? 'start' : 'end',
					vAlign: 'middle',
				};
			case 'ctr':
				return { x: x + w / 2, y: cy, anchor: 'middle', vAlign: 'middle' };
			default:
				return {
					x: endX + dir * LABEL_OFFSET,
					y: cy,
					anchor: positive ? 'start' : 'end',
					vAlign: 'middle',
				};
		}
	}
	const endY = positive ? y : y + h;
	const baseY = positive ? y + h : y;
	const cx = x + w / 2;
	switch (pos) {
		case 'inEnd':
			return positive
				? { x: cx, y: endY + LABEL_OFFSET, anchor: 'middle', vAlign: 'top' }
				: { x: cx, y: endY - LABEL_OFFSET, anchor: 'middle', vAlign: 'bottom' };
		case 'inBase':
			return positive
				? { x: cx, y: baseY - LABEL_OFFSET, anchor: 'middle', vAlign: 'bottom' }
				: { x: cx, y: baseY + LABEL_OFFSET, anchor: 'middle', vAlign: 'top' };
		case 'ctr':
			return { x: cx, y: y + h / 2, anchor: 'middle', vAlign: 'middle' };
		default:
			return positive
				? { x: cx, y: endY - LABEL_OFFSET, anchor: 'middle', vAlign: 'bottom' }
				: { x: cx, y: endY + LABEL_OFFSET, anchor: 'middle', vAlign: 'top' };
	}
}

/** Draw a bar plot; data labels are appended to `labels` (drawn above everything). */
export function drawBarPlot(
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
	const cross = crossValue(group);
	const stacks = stackValues(plot, plotData, cross);
	const visible = orderedSeries(plot).filter(({ index }) => !plotData.series[index]?.hidden);
	const stacked = isStacked(plot);
	const slot = slotFraction(group.cat);
	const geo = barGeometry(plot, slot, visible.length);
	const out: ChartRenderNode[] = [];
	visible.forEach(({ series, index }, k) => {
		const data = plotData.series[index];
		const points = stacks[index] ?? [];
		const nodes: ChartRenderNode[] = [];
		points.forEach((sv, i) => {
			if (sv.raw === null || !Number.isFinite(sv.raw) || (stacked && sv.from === sv.to)) {
				return;
			}
			const start = categoryFraction(group.cat, i) - slot / 2 + geo.lead + k * geo.step;
			const p0 = catPixel(group, inner, start);
			const p1 = catPixel(group, inner, start + geo.width);
			const v0 = valPixel(group, inner, sv.from);
			const v1 = valPixel(group, inner, sv.to);
			const rect: ChartBox = group.horizontal
				? {
						x: Math.min(v0, v1),
						y: Math.min(p0, p1),
						width: Math.abs(v1 - v0),
						height: Math.abs(p1 - p0),
					}
				: {
						x: Math.min(p0, p1),
						y: Math.min(v0, v1),
						width: Math.abs(p1 - p0),
						height: Math.abs(v1 - v0),
					};
			const paint = shapePaint(ctx, plot, series, i, points.length);
			const negative = sv.to < sv.from;
			const invert = findPoint(series, i)?.invertIfNegative ?? series.invertIfNegative;
			if (invert && negative) {
				paint.fill = '#FFFFFF';
				delete paint.gradient;
			}
			nodes.push({
				type: 'rect',
				...rect,
				...paint,
				role: 'bar',
				plot: plotIndex,
				series: index,
				point: i,
			});
			const lbl = effectiveLabels(plot, series, i);
			if (lbl) {
				const text = labelText(ctx, lbl, {
					value: sv.raw,
					category: data.categories[i] ?? String(i + 1),
					seriesName: data.name,
					percent: sv.share,
					valueFormat: data.valueFormat,
				});
				if (text) {
					// Value axes run upward (columns) or rightward (bars) unless reversed.
					const reversed = group.val.axis.orientation === 'maxMin';
					const a = barLabelAnchor(
						lbl.position,
						stacked,
						group.horizontal,
						rect,
						negative === reversed,
					);
					const style = labelTextStyle(ctx, plot, lbl);
					labels.push(
						...placeLabel(ctx, lbl, style, text, a.x, a.y, a.anchor, a.vAlign, {
							plot: plotIndex,
							series: index,
							point: i,
						}),
					);
				}
			}
		});
		out.push({ type: 'group', children: nodes, role: 'series', plot: plotIndex, series: index });
	});
	return out;
}

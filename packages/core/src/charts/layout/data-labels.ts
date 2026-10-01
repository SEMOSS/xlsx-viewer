/** Data label content, formatting and placement. */
import type { ChartDataLabels, ChartPlot, ChartSeries } from '../../types/chart';
import {
	type LayoutContext,
	type ResolvedText,
	fillPaint,
	formatNumber,
	resolveText,
	strokePaint,
} from './context';
import { blockHeight, textBlock } from './text-block';
import { blockWidth } from './text';
import type { ChartNodeMeta, ChartRenderNode, ChartTextNode } from './types';

/** Automatic data label color (Text 1 at 75%): #404040. */
const LABEL_COLOR = '#404040';

const FLAG_KEYS = [
	'showValue',
	'showCategoryName',
	'showSeriesName',
	'showPercent',
	'showBubbleSize',
] as const;

/**
 * Effective label settings for one point: the series labels (or the plot's),
 * overridden by the point's own `c:dLbl`. `undefined` when nothing is shown.
 */
export function effectiveLabels(
	plot: ChartPlot,
	series: ChartSeries,
	point: number,
): ChartDataLabels | undefined {
	const base = series.dataLabels ?? plot.dataLabels;
	if (!base) {
		return undefined;
	}
	const { points, ...rest } = base;
	const merged: ChartDataLabels = { ...rest };
	const own = points?.find((p) => p.idx === point);
	if (own) {
		for (const [k, v] of Object.entries(own)) {
			if (v !== undefined && k !== 'idx' && k !== 'layout') {
				(merged as Record<string, unknown>)[k] = v;
			}
		}
	}
	if (merged.deleted || !FLAG_KEYS.some((k) => merged[k])) {
		return undefined;
	}
	return merged;
}

export interface LabelInfo {
	value: number | null;
	category: string;
	seriesName: string;
	/** Share of the total in [0, 1] (pie charts, 100% stacks). */
	percent?: number;
	bubbleSize?: number | null;
	valueFormat: string;
}

/** Label text: series name, category, value, percentage and bubble size, in Excel's order. */
export function labelText(ctx: LayoutContext, labels: ChartDataLabels, info: LabelInfo): string {
	const parts: string[] = [];
	if (labels.showSeriesName) {
		parts.push(info.seriesName);
	}
	if (labels.showCategoryName) {
		parts.push(info.category);
	}
	if (labels.showValue && info.value !== null && Number.isFinite(info.value)) {
		parts.push(formatNumber(ctx, info.value, labels.numFmt ?? info.valueFormat));
	}
	if (labels.showPercent && info.percent !== undefined && Number.isFinite(info.percent)) {
		const code = labels.numFmt && !labels.showValue ? labels.numFmt : '0%';
		parts.push(formatNumber(ctx, info.percent, code));
	}
	if (labels.showBubbleSize && info.bubbleSize !== undefined && info.bubbleSize !== null) {
		parts.push(formatNumber(ctx, info.bubbleSize, 'General'));
	}
	return parts.join(labels.separator ?? ', ');
}

/** Text style of a label. */
export function labelTextStyle(
	ctx: LayoutContext,
	plot: ChartPlot,
	labels: ChartDataLabels,
): ResolvedText {
	return resolveText(ctx, { sizePt: 10, color: LABEL_COLOR }, plot.dataLabels?.style, labels.style);
}

/**
 * Label nodes for text anchored at (x, y): `anchor` aligns horizontally and
 * `vAlign` vertically relative to the point. Draws the label box when the
 * labels have a fill or border.
 */
export function placeLabel(
	ctx: LayoutContext,
	labels: ChartDataLabels,
	style: ResolvedText,
	text: string,
	x: number,
	y: number,
	anchor: ChartTextNode['anchor'],
	vAlign: 'top' | 'middle' | 'bottom',
	meta: ChartNodeMeta,
): ChartRenderNode[] {
	const lines = text.split('\n');
	const h = blockHeight(style, lines.length);
	const top = vAlign === 'top' ? y : vAlign === 'middle' ? y - h / 2 : y - h;
	const out: ChartRenderNode[] = [];
	const fill = fillPaint(ctx, labels.fill, null);
	const stroke = strokePaint(ctx, labels.line, null);
	if (fill.fill || fill.gradient || stroke.stroke) {
		const w = blockWidth(lines, style.font, ctx.measure);
		const left = anchor === 'start' ? x : anchor === 'end' ? x - w : x - w / 2;
		out.push({
			type: 'rect',
			x: left - 4,
			y: top - 2,
			width: w + 8,
			height: h + 4,
			...fill,
			...stroke,
			...meta,
			role: 'data-label',
		});
	}
	out.push(textBlock(style, lines, x, top, anchor, { ...meta, role: 'data-label' }));
	return out;
}

/** Size of a label's text block. */
export function labelSize(
	ctx: LayoutContext,
	style: ResolvedText,
	text: string,
): { width: number; height: number } {
	const lines = text.split('\n');
	return {
		width: blockWidth(lines, style.font, ctx.measure),
		height: blockHeight(style, lines.length),
	};
}

/**
 * Label for a point-like mark (line / scatter / radar marker, bubble) at
 * (x, y); `radius` is half the mark size. Returns no nodes when the point has
 * no label.
 */
export function pointLabel(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	point: number,
	x: number,
	y: number,
	radius: number,
	info: LabelInfo,
	defaultPosition: NonNullable<ChartDataLabels['position']>,
	meta: ChartNodeMeta,
): ChartRenderNode[] {
	const labels = effectiveLabels(plot, series, point);
	if (!labels) {
		return [];
	}
	const text = labelText(ctx, labels, info);
	if (!text) {
		return [];
	}
	const style = labelTextStyle(ctx, plot, labels);
	const gap = radius + 3;
	switch (labels.position ?? defaultPosition) {
		case 'r':
			return placeLabel(ctx, labels, style, text, x + gap, y, 'start', 'middle', meta);
		case 'l':
			return placeLabel(ctx, labels, style, text, x - gap, y, 'end', 'middle', meta);
		case 't':
			return placeLabel(ctx, labels, style, text, x, y - gap, 'middle', 'bottom', meta);
		case 'b':
			return placeLabel(ctx, labels, style, text, x, y + gap, 'middle', 'top', meta);
		default:
			return placeLabel(ctx, labels, style, text, x, y, 'middle', 'middle', meta);
	}
}

/**
 * Effective formatting of series and data points: explicit `c:spPr` /
 * `c:dPt` formatting, falling back to Excel's automatic colors, line widths
 * and markers.
 */
import type { ChartDataPoint, ChartMarker, ChartPlot, ChartSeries } from '../../types/chart';
import { type LayoutContext, PX_PER_PT, autoColor, fillPaint, strokePaint } from './context';
import { autoMarkerSymbol } from './palette';
import type { ChartPaint } from './types';

/** Automatic line width for line, scatter and radar series (2.25 pt). */
export const AUTO_LINE_WIDTH = 2.25 * PX_PER_PT;

const PIE_FAMILY = new Set(['pie', 'doughnut', 'ofPie']);

/** Whether a plot colors each data point differently. */
export function variesByPoint(plot: ChartPlot): boolean {
	if (PIE_FAMILY.has(plot.type)) {
		return plot.varyColors ?? true;
	}
	return Boolean(plot.varyColors) && plot.series.length === 1;
}

/** Automatic color of a series (by its `c:idx`) or of a point when colors vary by point. */
export function seriesAutoColor(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	point?: number,
	pointCount = 0,
): string {
	if (point !== undefined && variesByPoint(plot)) {
		return autoColor(ctx, point, pointCount);
	}
	const count = ctx.chart.plots.reduce((n, p) => n + p.series.length, 0);
	return autoColor(ctx, series.idx, count);
}

export function findPoint(series: ChartSeries, idx: number): ChartDataPoint | undefined {
	return series.points?.find((p) => p.idx === idx);
}

/** Fill + outline for a filled shape (bar, area, slice, bubble) of a series or point. */
export function shapePaint(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	point?: number,
	pointCount = 0,
	outline: { color: string; width: number } | null = null,
): ChartPaint {
	const auto = seriesAutoColor(ctx, plot, series, point, pointCount);
	const dp = point === undefined ? undefined : findPoint(series, point);
	const varies = point !== undefined && variesByPoint(plot);
	// With varying colors the series fill does not apply to points without their own fill.
	const fill = dp?.fill ?? (varies ? undefined : series.fill);
	const line = dp?.line ?? series.line;
	return { ...fillPaint(ctx, fill, auto), ...strokePaint(ctx, line, outline) };
}

/** Stroke for a line-type series (line, scatter, radar, stock), or none. */
export function linePaint(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	drawByDefault = true,
): ChartPaint {
	const auto = seriesAutoColor(ctx, plot, series);
	return strokePaint(
		ctx,
		series.line,
		drawByDefault ? { color: auto, width: AUTO_LINE_WIDTH, cap: 'round' } : null,
	);
}

export interface ResolvedMarker {
	symbol: NonNullable<ChartMarker['symbol']>;
	/** Box size in pixels. */
	size: number;
	paint: ChartPaint;
}

/** Whether a plot shows automatic markers on series that do not specify one. */
function autoMarkers(plot: ChartPlot): boolean {
	switch (plot.type) {
		case 'line':
			return plot.marker !== false && !plot.is3D;
		case 'scatter':
			return (
				plot.scatterStyle !== 'line' &&
				plot.scatterStyle !== 'smooth' &&
				plot.scatterStyle !== 'none'
			);
		case 'radar':
			return plot.radarStyle === 'marker';
		default:
			return false;
	}
}

/** The marker drawn at a point of a series, or `undefined` for none. */
export function resolveMarker(
	ctx: LayoutContext,
	plot: ChartPlot,
	series: ChartSeries,
	point?: number,
): ResolvedMarker | undefined {
	const dp = point === undefined ? undefined : findPoint(series, point);
	const m: ChartMarker = { ...(series.marker ?? {}), ...(dp?.marker ?? {}) };
	let symbol = m.symbol;
	let auto = false;
	if (symbol === 'none') {
		return undefined;
	}
	if (!symbol || symbol === 'auto') {
		if (!autoMarkers(plot)) {
			return undefined;
		}
		symbol = autoMarkerSymbol(series.idx);
		auto = !m.symbol;
	}
	const sizePt = m.size ?? (auto ? 7 : 5);
	const color = seriesAutoColor(ctx, plot, series, point);
	const lineSymbol = symbol === 'x' || symbol === 'plus' || symbol === 'star';
	const fill = fillPaint(ctx, m.fill, lineSymbol ? null : color);
	const stroke = strokePaint(ctx, m.line, { color, width: 0.75 * PX_PER_PT });
	return { symbol, size: Math.max(2, sizePt * PX_PER_PT), paint: { ...fill, ...stroke } };
}

/** Series sorted by plot order (`c:order`), keeping their index in `plot.series`. */
export function orderedSeries(plot: ChartPlot): Array<{ series: ChartSeries; index: number }> {
	return plot.series
		.map((series, index) => ({ series, index }))
		.sort((a, b) => a.series.order - b.series.order || a.index - b.index);
}

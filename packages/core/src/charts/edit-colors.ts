/**
 * Change Colors and Chart Styles: write explicit series / point colors from
 * Excel's color palettes, and apply a small set of quick styles.
 */
import type { ChartAxis, ChartModel, ChartPlot, ChartSeries } from '../types/chart';
import type { DrawingColor, DrawingTextStyle, ShapeFill, ShapeLine } from '../types/drawing';
import { GRID_COLOR, NO_LINE, TEXT_COLOR, hairline, solid } from './data/chart-defaults';
import { cloneChart, familyOfPlot } from './edit-series';
import { accentVariation, monochromeColor } from './layout/palette';
import { variesByPoint } from './layout/series-style';

export type ChartColorPalette =
	| 'colorful1'
	| 'colorful2'
	| 'colorful3'
	| 'colorful4'
	| 'monochromatic1'
	| 'monochromatic2'
	| 'monochromatic3'
	| 'monochromatic4'
	| 'monochromatic5'
	| 'monochromatic6';

/** Every palette, in gallery order. */
export const CHART_COLOR_PALETTES: readonly ChartColorPalette[] = [
	'colorful1',
	'colorful2',
	'colorful3',
	'colorful4',
	'monochromatic1',
	'monochromatic2',
	'monochromatic3',
	'monochromatic4',
	'monochromatic5',
	'monochromatic6',
];

/** Accent cycles of the colorful palettes (Office color styles 10-13). */
const COLORFUL: Record<string, number[]> = {
	colorful1: [1, 2, 3, 4, 5, 6],
	colorful2: [1, 3, 5],
	colorful3: [2, 4, 6],
	colorful4: [6, 5, 4, 3, 2, 1],
};

/** Color `index` of `count` in a palette. */
export function chartPaletteColor(
	palette: ChartColorPalette,
	index: number,
	count: number,
): DrawingColor {
	const mono = /^monochromatic(\d)$/.exec(palette);
	if (mono) {
		return monochromeColor(`accent${mono[1]}`, index, count);
	}
	const accents = COLORFUL[palette] ?? COLORFUL.colorful1;
	const i = Math.max(0, index);
	return accentVariation(`accent${accents[i % accents.length]}`, Math.floor(i / accents.length));
}

function pointCount(series: ChartSeries): number {
	return Math.max(series.values?.cache?.length ?? 0, series.categories?.cache?.length ?? 0);
}

function isFilled(plot: ChartPlot): boolean {
	const family = familyOfPlot(plot);
	return (
		family === 'bar' ||
		family === 'area' ||
		family === 'pie' ||
		plot.type === 'bubble' ||
		plot.type === 'surface' ||
		(plot.type === 'radar' && plot.radarStyle === 'filled')
	);
}

function recolorSeries(series: ChartSeries, plot: ChartPlot, color: DrawingColor): void {
	if (isFilled(plot)) {
		series.fill = solid(color);
		return;
	}
	if (series.line?.fill?.type !== 'none') {
		series.line = { ...(series.line ?? {}), fill: solid(color) };
	}
	if (series.marker?.symbol !== 'none') {
		series.marker = {
			...(series.marker ?? {}),
			fill: solid(color),
			line: { ...(series.marker?.line ?? { width: 9525 }), fill: solid(color) },
		};
	}
}

/**
 * Excel's Change Colors: every series (or every point of a vary-colors chart
 * such as a pie) gets an explicit color from `palette`.
 */
export function applyColorPalette(model: ChartModel, palette: ChartColorPalette): ChartModel {
	const next = cloneChart(model);
	const total = next.plots.reduce((n, p) => n + p.series.length, 0);
	let index = 0;
	for (const plot of next.plots) {
		const byPoint = variesByPoint(plot);
		for (const series of [...plot.series].sort((a, b) => a.order - b.order)) {
			if (byPoint) {
				const count = pointCount(series);
				const points = new Map((series.points ?? []).map((p) => [p.idx, p] as const));
				for (let i = 0; i < count; i++) {
					points.set(i, {
						...(points.get(i) ?? { idx: i }),
						fill: solid(chartPaletteColor(palette, i, count)),
					});
				}
				series.points = [...points.values()].sort((a, b) => a.idx - b.idx);
			} else {
				recolorSeries(series, plot, chartPaletteColor(palette, index, total));
			}
			index++;
		}
	}
	next.dirty = true;
	return next;
}

interface QuickStyle {
	chartFill: ShapeFill;
	chartLine: ShapeLine;
	plotFill: ShapeFill;
	/** Primary value-axis major gridlines; `null` removes them. */
	gridlines: ShapeLine | null;
	/** Category axis line. */
	axisLine: ShapeLine;
	text: DrawingColor;
	/** Show value data labels. */
	labels: boolean;
	/** Outline of filled series (bars, areas, slices). */
	outline?: ShapeLine;
	gapWidth?: number;
}

const tx1 = (lumMod: number, lumOff: number): DrawingColor => ({
	kind: 'scheme',
	value: 'tx1',
	transforms: [
		{ name: 'lumMod', val: lumMod },
		{ name: 'lumOff', val: lumOff },
	],
});
const bg1 = (lumMod?: number): DrawingColor =>
	lumMod === undefined
		? { kind: 'scheme', value: 'bg1' }
		: { kind: 'scheme', value: 'bg1', transforms: [{ name: 'lumMod', val: lumMod }] };

/** Quick styles 1-6, approximating the first entries of Excel's Chart Styles gallery. */
const QUICK_STYLES: Record<number, QuickStyle> = {
	1: {
		chartFill: solid(bg1()),
		chartLine: hairline(GRID_COLOR),
		plotFill: { type: 'none' },
		gridlines: hairline(GRID_COLOR),
		axisLine: hairline(GRID_COLOR),
		text: TEXT_COLOR,
		labels: false,
	},
	2: {
		chartFill: solid(bg1()),
		chartLine: hairline(GRID_COLOR),
		plotFill: { type: 'none' },
		gridlines: null,
		axisLine: hairline(GRID_COLOR),
		text: TEXT_COLOR,
		labels: true,
		gapWidth: 100,
	},
	3: {
		chartFill: solid(bg1()),
		chartLine: hairline(GRID_COLOR),
		plotFill: { type: 'none' },
		gridlines: { ...hairline(GRID_COLOR), dash: 'dash' },
		axisLine: hairline(tx1(25000, 75000)),
		text: TEXT_COLOR,
		labels: false,
		gapWidth: 150,
	},
	4: {
		chartFill: solid(bg1()),
		chartLine: hairline(GRID_COLOR),
		plotFill: { type: 'none' },
		gridlines: hairline(GRID_COLOR),
		axisLine: hairline(GRID_COLOR),
		text: TEXT_COLOR,
		labels: false,
		outline: { width: 19050, fill: solid(bg1()) },
		gapWidth: 50,
	},
	5: {
		chartFill: solid(bg1(95000)),
		chartLine: NO_LINE,
		plotFill: { type: 'none' },
		gridlines: { width: 9525, fill: solid(bg1()) },
		axisLine: hairline(bg1()),
		text: TEXT_COLOR,
		labels: false,
	},
	6: {
		chartFill: solid(tx1(75000, 25000)),
		chartLine: NO_LINE,
		plotFill: { type: 'none' },
		gridlines: hairline(tx1(50000, 50000)),
		axisLine: hairline(tx1(50000, 50000)),
		text: bg1(85000),
		labels: false,
	},
};

/** Number of quick styles `applyChartStyle` offers. */
export const CHART_QUICK_STYLE_COUNT = 6;

const recolorText = (
	style: DrawingTextStyle | undefined,
	color: DrawingColor,
): DrawingTextStyle => ({ ...(style ?? {}), color });

function primaryValueAxis(model: ChartModel): ChartAxis | undefined {
	const plot = model.plots.find((p) => p.axisIds.length >= 2);
	const axes = (plot?.axisIds ?? []).map((id) => model.axes.find((a) => a.id === id));
	const vals = axes.filter((a): a is ChartAxis => a?.kind === 'val');
	return vals.find((a) => a.position === 'l' || a.position === 'r') ?? vals[0];
}

/**
 * Apply quick style `styleIndex` (1-6): chart and plot fills, gridlines,
 * axis lines, text color, data labels and bar outlines / gap width.
 */
export function applyChartStyle(model: ChartModel, styleIndex: number): ChartModel {
	const style =
		QUICK_STYLES[Math.min(CHART_QUICK_STYLE_COUNT, Math.max(1, Math.round(styleIndex)))];
	const next = cloneChart(model);
	next.chartFill = style.chartFill;
	next.chartLine = style.chartLine;
	next.plotFill = style.plotFill;
	next.textStyle = recolorText(next.textStyle, style.text);
	if (next.title) {
		next.title.style = recolorText(next.title.style, style.text);
	}
	if (next.legend) {
		next.legend.textStyle = recolorText(next.legend.textStyle, style.text);
	}
	const valueAxis = primaryValueAxis(next);
	for (const axis of next.axes) {
		axis.textStyle = recolorText(axis.textStyle, style.text);
		if (axis.title) {
			axis.title.style = recolorText(axis.title.style, style.text);
		}
		if (axis === valueAxis) {
			if (style.gridlines) {
				axis.majorGridlines = style.gridlines;
			} else {
				delete axis.majorGridlines;
			}
		} else if (axis.kind !== 'val' && axis.kind !== 'ser') {
			axis.line = style.axisLine;
		}
	}
	for (const plot of next.plots) {
		const pie = familyOfPlot(plot) === 'pie';
		if (plot.type === 'bar') {
			const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
			plot.gapWidth = style.gapWidth ?? (stacked ? 150 : 219);
		}
		for (const series of plot.series) {
			if (isFilled(plot) && !pie) {
				series.line = style.outline ?? NO_LINE;
			}
			if (style.labels) {
				series.dataLabels = {
					...(series.dataLabels ?? {}),
					showValue: !pie,
					showPercent: pie,
					showCategoryName: false,
					showSeriesName: false,
					showLegendKey: false,
					style: recolorText(series.dataLabels?.style, style.text),
				};
			} else {
				delete series.dataLabels;
			}
		}
		delete plot.dataLabels;
	}
	next.dirty = true;
	return next;
}

/**
 * Series formatting for Change Chart Type: carry a series' color and data over
 * to the fill, outline and markers the new chart type needs.
 */
import type { ChartDataLabels, ChartPlot, ChartSeries } from '../types/chart';
import type { DrawingColor, ShapeFill, ShapeLine } from '../types/drawing';
import { NO_LINE, solid } from './data/chart-defaults';

export type ChartTypeTarget =
	| 'columnClustered'
	| 'columnStacked'
	| 'columnPercent'
	| 'barClustered'
	| 'barStacked'
	| 'barPercent'
	| 'line'
	| 'lineMarkers'
	| 'area'
	| 'areaStacked'
	| 'pie'
	| 'doughnut'
	| 'scatter'
	| 'scatterLines'
	| 'radar';

/** Every conversion target, in gallery order. */
export const CHART_TYPE_TARGETS: readonly ChartTypeTarget[] = [
	'columnClustered',
	'columnStacked',
	'columnPercent',
	'barClustered',
	'barStacked',
	'barPercent',
	'line',
	'lineMarkers',
	'area',
	'areaStacked',
	'pie',
	'doughnut',
	'scatter',
	'scatterLines',
	'radar',
];

export type ChartFamily = 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar' | 'other';

export function familyOfPlot(plot: ChartPlot): ChartFamily {
	switch (plot.type) {
		case 'bar':
		case 'line':
		case 'area':
		case 'scatter':
		case 'radar':
			return plot.type;
		case 'pie':
		case 'doughnut':
		case 'ofPie':
			return 'pie';
		case 'stock':
			return 'line';
		case 'bubble':
			return 'scatter';
		default:
			return 'other';
	}
}

export function familyOfTarget(target: ChartTypeTarget): ChartFamily {
	if (target.startsWith('column') || target.startsWith('bar')) {
		return 'bar';
	}
	if (target === 'doughnut') {
		return 'pie';
	}
	return target.replace(/(Markers|Stacked|Lines)$/, '') as ChartFamily;
}

/** Deep copy of plain chart data (models never hold functions or cycles). */
export function cloneChart<T>(value: T): T {
	return typeof structuredClone === 'function'
		? structuredClone(value)
		: (JSON.parse(JSON.stringify(value)) as T);
}

const solidColor = (fill: ShapeFill | undefined): DrawingColor | undefined =>
	fill?.type === 'solid' ? fill.color : undefined;

/** The color that identifies a series (fill for filled plots, line or marker otherwise). */
export function seriesColor(series: ChartSeries, plot: ChartPlot): DrawingColor | undefined {
	const family = familyOfPlot(plot);
	const filled =
		family === 'bar' ||
		family === 'area' ||
		family === 'pie' ||
		plot.type === 'bubble' ||
		(family === 'radar' && plot.radarStyle === 'filled');
	if (filled) {
		return solidColor(series.fill) ?? solidColor(series.line?.fill);
	}
	return (
		solidColor(series.line?.fill) ?? solidColor(series.marker?.fill) ?? solidColor(series.fill)
	);
}

/** Label positions each target allows (an invalid position makes Excel repair the file). */
export function allowedPositions(target: ChartTypeTarget): readonly string[] {
	if (target === 'columnClustered' || target === 'barClustered') {
		return ['ctr', 'inEnd', 'inBase', 'outEnd'];
	}
	if (familyOfTarget(target) === 'bar') {
		return ['ctr', 'inEnd', 'inBase'];
	}
	if (target === 'pie') {
		return ['bestFit', 'ctr', 'inEnd', 'outEnd'];
	}
	if (target.startsWith('line') || target.startsWith('scatter')) {
		return ['ctr', 'l', 'r', 't', 'b'];
	}
	return [];
}

export function fixLabels(
	labels: ChartDataLabels | undefined,
	allowed: readonly string[],
): ChartDataLabels | undefined {
	if (!labels) {
		return undefined;
	}
	const fix = <T extends ChartDataLabels>(l: T): T => {
		const out = { ...l };
		if (out.position && !allowed.includes(out.position)) {
			delete out.position;
		}
		return out;
	};
	const out = fix(labels);
	if (labels.points) {
		out.points = labels.points.map(fix);
	}
	return out;
}

const visible = (line: ShapeLine | undefined): boolean =>
	Boolean(line) && line!.fill?.type !== 'none';

/** Markers for a type that shows them: keep visible ones, else Excel's circle in the series color. */
function markerFor(
	color: DrawingColor | undefined,
	source: ChartSeries,
): NonNullable<ChartSeries['marker']> {
	const from = source.marker;
	if (from?.symbol && from.symbol !== 'none' && from.symbol !== 'auto') {
		return cloneChart(from);
	}
	const m: NonNullable<ChartSeries['marker']> = { symbol: 'circle', size: from?.size ?? 5 };
	if (color) {
		m.fill = solid(color);
		m.line = { width: 9525, fill: solid(color) };
	}
	return m;
}

/**
 * Series formatting for a target type: the series keeps its color (`color`)
 * and data, and gets the fill, outline and markers the new type needs.
 */
export function restyleSeries(
	source: ChartSeries,
	sourcePlot: ChartPlot,
	target: ChartTypeTarget,
	color: DrawingColor | undefined,
): ChartSeries {
	const from = familyOfPlot(sourcePlot);
	const to = familyOfTarget(target);
	const s: ChartSeries = { idx: source.idx, order: source.order };
	for (const key of ['name', 'categories', 'values'] as const) {
		if (source[key]) {
			s[key] = cloneChart(source[key]) as never;
		}
	}
	const labels = fixLabels(source.dataLabels, allowedPositions(target));
	if (labels) {
		s.dataLabels = labels;
	}
	if (from === to && source.points?.length) {
		s.points = cloneChart(source.points);
	}
	const trendlines = to === 'bar' || to === 'line' || to === 'area' || to === 'scatter';
	const stacked = /Stacked|Percent/.test(target);
	if (trendlines && !stacked && source.trendlines?.length) {
		s.trendlines = cloneChart(source.trendlines);
	}
	const line = (width: number): ShapeLine => {
		const l: ShapeLine = { width, cap: 'rnd' };
		if (color) {
			l.fill = solid(color);
		}
		return l;
	};
	switch (to) {
		case 'bar':
		case 'area':
		case 'pie':
			// Pies vary colors by point, so a series color only carries over between pies.
			if (to !== 'pie' && color) {
				s.fill = solid(color);
			} else if (to === 'pie' && from === 'pie' && source.fill) {
				s.fill = cloneChart(source.fill);
			}
			if (to === 'pie' && from !== 'pie') {
				s.line = { width: 19050, fill: solid({ kind: 'scheme', value: 'bg1' }) };
			} else if (source.line && (from === to || from === 'bar' || from === 'area')) {
				s.line = cloneChart(source.line);
			} else if (color && to !== 'pie') {
				s.line = NO_LINE;
			}
			if (to === 'bar' && from === 'bar' && source.invertIfNegative !== undefined) {
				s.invertIfNegative = source.invertIfNegative;
			}
			if (to === 'pie' && from === 'pie' && source.explosion !== undefined) {
				s.explosion = source.explosion;
			}
			break;
		case 'line':
			s.line = from === 'line' && visible(source.line) ? cloneChart(source.line!) : line(28575);
			s.marker = target === 'lineMarkers' ? markerFor(color, source) : { symbol: 'none' };
			s.smooth = from === 'line' || from === 'scatter' ? Boolean(source.smooth) : false;
			break;
		case 'scatter':
			if (target === 'scatter') {
				s.line = { width: 25400, cap: 'rnd', fill: { type: 'none' } };
			} else {
				s.line =
					(from === 'scatter' || from === 'line') && visible(source.line)
						? cloneChart(source.line!)
						: line(19050);
			}
			s.marker = markerFor(color, source);
			s.smooth = false;
			break;
		default:
			s.line = from === 'radar' && visible(source.line) ? cloneChart(source.line!) : line(28575);
			s.marker = { symbol: 'none' };
	}
	return s;
}

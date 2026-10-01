/**
 * Change Chart Type: convert a chart (all its plots) to one chart type the way
 * Excel does, keeping series references, names, colors, the title and the
 * axes that still apply.
 */
import type { ChartAxis, ChartModel, ChartPlot, ChartSeries } from '../types/chart';
import { defaultCategoryAxis, defaultValueAxis } from './data/chart-defaults';
import {
	type ChartTypeTarget,
	allowedPositions,
	cloneChart,
	familyOfTarget,
	fixLabels,
	restyleSeries,
	seriesColor,
} from './edit-series';

export { CHART_TYPE_TARGETS, cloneChart } from './edit-series';
export type { ChartTypeTarget } from './edit-series';

/** The primary axes of a chart: the category (X) axis and the value (Y) axis. */
function primaryAxes(model: ChartModel): { cat?: ChartAxis; val?: ChartAxis } {
	const plot = model.plots.find((p) => p.axisIds.length >= 2);
	if (!plot) {
		return {};
	}
	let a = model.axes.find((x) => x.id === plot.axisIds[0]);
	let b = model.axes.find((x) => x.id === plot.axisIds[1]);
	if (a?.kind === 'val' && b?.kind === 'val' && (a.position === 'l' || a.position === 'r')) {
		[a, b] = [b, a];
	} else if (a?.kind === 'val' && b && b.kind !== 'val') {
		[a, b] = [b, a];
	}
	return { cat: a, val: b?.kind === 'val' ? b : undefined };
}

const VALUE_ONLY = [
	'min',
	'max',
	'majorUnit',
	'minorUnit',
	'logBase',
	'displayUnits',
	'crossBetween',
] as const;
const CATEGORY_ONLY = [
	'tickLabelSkip',
	'noMultiLevelLabels',
	'baseTimeUnit',
	'majorTimeUnit',
] as const;

/** Axes for a target: reuse the primary pair (retyped and repositioned) or create Excel's defaults. */
function axesFor(model: ChartModel, target: ChartTypeTarget): ChartAxis[] {
	const family = familyOfTarget(target);
	if (family === 'pie') {
		return [];
	}
	const horizontal = target.startsWith('bar');
	const xy = family === 'scatter';
	const between = family === 'area' || xy ? 'midCat' : 'between';
	const found = primaryAxes(model);
	let cat = found.cat ? cloneChart(found.cat) : undefined;
	let val = found.val ? cloneChart(found.val) : undefined;
	if (!cat || !val) {
		const base = Math.max(500000000, ...model.axes.map((a) => a.id + 1));
		cat = defaultCategoryAxis(base, base + 1, horizontal ? 'l' : 'b', xy ? 'val' : 'cat');
		val = defaultValueAxis(base + 1, base, horizontal ? 'b' : 'l', between);
		return [cat, val];
	}
	if (xy && cat.kind !== 'val') {
		for (const key of CATEGORY_ONLY) {
			delete cat[key];
		}
		cat.kind = 'val';
		cat.crossBetween = 'midCat';
	} else if (!xy && cat.kind === 'val') {
		for (const key of VALUE_ONLY) {
			delete cat[key];
		}
		cat.kind = 'cat';
	} else if (!xy && cat.kind === 'ser') {
		cat.kind = 'cat';
	}
	const far = (a: ChartAxis) => a.position === 't' || a.position === 'r';
	cat.position = horizontal ? (far(cat) ? 'r' : 'l') : far(cat) ? 't' : 'b';
	val.position = horizontal ? (far(val) ? 't' : 'b') : far(val) ? 'r' : 'l';
	cat.crossAxisId = val.id;
	val.crossAxisId = cat.id;
	val.crossBetween = between;
	return [cat, val];
}

function plotFor(
	target: ChartTypeTarget,
	previous: ChartPlot | undefined,
): Omit<ChartPlot, 'series' | 'axisIds'> {
	const stacked = target.endsWith('Stacked')
		? 'stacked'
		: target.endsWith('Percent')
			? 'percentStacked'
			: undefined;
	switch (familyOfTarget(target)) {
		case 'bar': {
			const plot: Omit<ChartPlot, 'series' | 'axisIds'> = {
				type: 'bar',
				barDirection: target.startsWith('bar') ? 'bar' : 'col',
				grouping: stacked ?? 'clustered',
				varyColors: false,
				gapWidth:
					previous?.type === 'bar' && previous.gapWidth !== undefined
						? previous.gapWidth
						: stacked
							? 150
							: 219,
				overlap: stacked
					? 100
					: previous?.type === 'bar' &&
						  previous.grouping === 'clustered' &&
						  previous.overlap !== undefined
						? previous.overlap
						: -27,
			};
			return plot;
		}
		case 'line':
			return { type: 'line', grouping: 'standard', varyColors: false, marker: true };
		case 'area':
			return { type: 'area', grouping: stacked ?? 'standard', varyColors: false };
		case 'pie': {
			const firstSliceAngle = previous?.firstSliceAngle ?? 0;
			return target === 'doughnut'
				? {
						type: 'doughnut',
						varyColors: true,
						firstSliceAngle,
						holeSize: previous?.holeSize ?? 75,
					}
				: { type: 'pie', varyColors: true, firstSliceAngle };
		}
		case 'scatter':
			return { type: 'scatter', scatterStyle: 'lineMarker', varyColors: false };
		default:
			return { type: 'radar', radarStyle: 'marker', varyColors: false };
	}
}

/**
 * Convert every plot of a chart to `target`. Series keep their references,
 * names and colors; axes are reused, retyped or created as the type needs
 * (pie and doughnut charts have none, scatter charts have two value axes).
 */
export function convertChartType(model: ChartModel, target: ChartTypeTarget): ChartModel {
	const series: ChartSeries[] = [];
	for (const plot of model.plots) {
		const ordered = [...plot.series].sort((a, b) => a.order - b.order);
		for (const s of ordered) {
			series.push(restyleSeries(s, plot, target, seriesColor(s, plot)));
		}
	}
	const axes = axesFor(model, target);
	const previous = model.plots[0];
	const plot: ChartPlot = { ...plotFor(target, previous), series, axisIds: axes.map((a) => a.id) };
	const labels = fixLabels(previous?.dataLabels, allowedPositions(target));
	if (labels) {
		plot.dataLabels = labels;
	}
	return { ...cloneChart(model), plots: [plot], axes, dirty: true };
}

/** The gallery entry matching a chart's (first) plot, if it is one of the targets. */
export function chartTypeTargetOf(model: ChartModel): ChartTypeTarget | undefined {
	const plot = model.plots[0];
	if (!plot) {
		return undefined;
	}
	const suffix =
		plot.grouping === 'stacked'
			? 'Stacked'
			: plot.grouping === 'percentStacked'
				? 'Percent'
				: 'Clustered';
	switch (plot.type) {
		case 'bar':
			return `${plot.barDirection === 'bar' ? 'bar' : 'column'}${suffix}` as ChartTypeTarget;
		case 'line': {
			const markers = plot.marker !== false && plot.series.some((s) => s.marker?.symbol !== 'none');
			return markers ? 'lineMarkers' : 'line';
		}
		case 'area':
			return plot.grouping === 'stacked' || plot.grouping === 'percentStacked'
				? 'areaStacked'
				: 'area';
		case 'pie':
		case 'ofPie':
			return 'pie';
		case 'doughnut':
			return 'doughnut';
		case 'scatter':
			return plot.series.every((s) => s.line?.fill?.type === 'none') ? 'scatter' : 'scatterLines';
		case 'radar':
			return 'radar';
		default:
			return undefined;
	}
}

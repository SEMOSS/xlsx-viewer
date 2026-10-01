/** Plain-language chart summaries (accessibility labels). */
import type { ChartModel, ChartPlot } from '../../types/chart';
import type { ChartData } from '../data/types';

function plotNoun(plot: ChartPlot): string {
	const grouping =
		plot.grouping === 'stacked'
			? 'stacked '
			: plot.grouping === 'percentStacked'
				? '100% stacked '
				: '';
	switch (plot.type) {
		case 'bar':
			return `${grouping}${plot.barDirection === 'bar' ? 'bar' : 'column'}`;
		case 'line':
		case 'area':
			return `${grouping}${plot.type}`;
		case 'ofPie':
			return 'pie';
		case 'scatter':
			return 'scatter';
		case 'unknown':
			return 'unsupported';
		default:
			return plot.type;
	}
}

/** The title Excel shows: explicit text, else the series name of a single-series chart, else "Chart Title". */
export function chartTitleText(chart: ChartModel, data?: ChartData): string | undefined {
	if (!chart.title) {
		return undefined;
	}
	const linked = data?.titles?.chart;
	if (linked !== undefined) {
		return linked;
	}
	if (chart.title.text !== undefined) {
		return chart.title.text;
	}
	// A pie draws only its first series, so that series names the chart.
	const pieOnly =
		chart.plots.length === 1 && (chart.plots[0].type === 'pie' || chart.plots[0].type === 'ofPie');
	const visible = (data?.plots ?? []).flatMap((p) => p.series.filter((s) => !s.hidden));
	const modelSeries = chart.plots.flatMap((p) => p.series);
	if (visible.length === 1 || (pieOnly && visible.length > 0)) {
		return visible[0].name;
	}
	if (!data && (modelSeries.length === 1 || (pieOnly && modelSeries.length > 0))) {
		return modelSeries[0].name?.text ?? 'Chart Title';
	}
	return 'Chart Title';
}

/** Summary such as `Sales, column chart with 2 series`. */
export function describeChart(chart: ChartModel, data?: ChartData): string {
	const title = chartTitleText(chart, data);
	const kinds = [...new Set(chart.plots.map(plotNoun))];
	const kind =
		kinds.length === 0 ? 'chart' : kinds.length === 1 ? `${kinds[0]} chart` : 'combo chart';
	const seriesCount = chart.plots.reduce((n, p) => n + p.series.length, 0);
	const parts = [title, `${kind} with ${seriesCount} series`].filter((p): p is string =>
		Boolean(p),
	);
	return parts.join(', ');
}

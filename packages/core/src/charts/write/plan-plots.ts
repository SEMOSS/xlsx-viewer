/**
 * Pair chart plots with the container element they are written as and the
 * axes they reference, creating default axes for cartesian plots without any.
 */
import type { ChartAxis, ChartModel, ChartPlot } from '../../types/chart';

const PIE_FAMILY = new Set(['pie', 'doughnut', 'ofPie']);

/** The container element for a plot, given how many valid axes it references. */
function containerName(plot: ChartPlot, axisCount: number): string | undefined {
	switch (plot.type) {
		case 'bar':
			return plot.is3D ? 'bar3DChart' : 'barChart';
		case 'line':
			return plot.is3D && axisCount >= 3 ? 'line3DChart' : 'lineChart';
		case 'pie':
			return plot.is3D ? 'pie3DChart' : 'pieChart';
		case 'doughnut':
			return 'doughnutChart';
		case 'ofPie':
			return 'ofPieChart';
		case 'area':
			return plot.is3D ? 'area3DChart' : 'areaChart';
		case 'scatter':
			return 'scatterChart';
		case 'radar':
			return 'radarChart';
		case 'bubble':
			return 'bubbleChart';
		case 'stock':
			return plot.series.length >= 3 && plot.series.length <= 4 ? 'stockChart' : 'lineChart';
		case 'surface':
			return plot.is3D && axisCount >= 3 ? 'surface3DChart' : 'surfaceChart';
		default:
			return undefined;
	}
}

/** Containers that accept a third (series) axis id. */
const THREE_AXES = new Set([
	'bar3DChart',
	'line3DChart',
	'area3DChart',
	'surfaceChart',
	'surface3DChart',
]);

export interface PlannedPlot {
	plot: ChartPlot;
	container: string;
	axisIds: number[];
}

/** Pair every plot with valid axes, creating default axes for plots that lack them. */
export function planPlots(chart: ChartModel): { plots: PlannedPlot[]; axes: ChartAxis[] } {
	const byId = new Map(chart.axes.map((a) => [a.id, a] as const));
	let nextId = Math.max(500000000, ...chart.axes.map((a) => a.id + 1));
	const created: ChartAxis[] = [];
	const plots: PlannedPlot[] = [];
	for (const plot of chart.plots) {
		if (PIE_FAMILY.has(plot.type)) {
			const container = containerName(plot, 0);
			if (container) {
				plots.push({ plot, container, axisIds: [] });
			}
			continue;
		}
		let ids = plot.axisIds.filter((id) => byId.has(id));
		if (ids.length < 2) {
			const xy = plot.type === 'scatter' || plot.type === 'bubble';
			const a: ChartAxis = {
				id: nextId++,
				kind: xy ? 'val' : 'cat',
				position: 'b',
				crossAxisId: 0,
			};
			const b: ChartAxis = { id: nextId++, kind: 'val', position: 'l', crossAxisId: a.id };
			a.crossAxisId = b.id;
			if (plot.type === 'bar' && plot.barDirection === 'bar') {
				a.position = 'l';
				b.position = 'b';
			}
			for (const axis of [a, b]) {
				byId.set(axis.id, axis);
				created.push(axis);
			}
			ids = [a.id, b.id];
		}
		const container = containerName(plot, ids.length);
		if (!container) {
			continue;
		}
		plots.push({ plot, container, axisIds: ids.slice(0, THREE_AXES.has(container) ? 3 : 2) });
	}
	const referenced = new Set(plots.flatMap((p) => p.axisIds));
	const axes = [...chart.axes, ...created].filter((a) => referenced.has(a.id));
	return { plots, axes };
}

/**
 * Chart editing helpers behind Excel's Chart Design tab. Every helper is pure:
 * it returns a new ChartModel (marked `dirty`) and never mutates its input.
 * Helpers that cannot apply return the original model (or `null`).
 */
import type { RangeAddress } from '../types/address';
import type { ChartAxis, ChartLegend, ChartModel, ChartPlot } from '../types/chart';
import type { Workbook } from '../types/workbook';
import { GRID_COLOR, TEXT_COLOR, hairline, labelText } from './data/chart-defaults';
import { resolveChartRef } from './data/chart-ref';
import { type CreatableChartType, createChartFromRange } from './data/create-chart';
import { cloneChart, familyOfPlot, restyleSeries } from './edit-series';
import { chartTypeTargetOf } from './edit-type';
import { autoChartColor } from './layout/palette';

export {
	applyChartStyle,
	applyColorPalette,
	chartPaletteColor,
	CHART_COLOR_PALETTES,
	CHART_QUICK_STYLE_COUNT,
} from './edit-colors';
export type { ChartColorPalette } from './edit-colors';
export { CHART_TYPE_TARGETS, chartTypeTargetOf, convertChartType } from './edit-type';
export type { ChartTypeTarget } from './edit-type';

/** Options of each element in Excel's Add Chart Element menu. */
export interface ChartElementOptions {
	chartTitle: 'none' | 'above' | 'overlay';
	legend: 'none' | 'right' | 'top' | 'left' | 'bottom';
	dataLabels: 'none' | 'show';
	/** Primary major gridlines of the value axis. */
	gridlines: 'on' | 'off';
	/** Titles of the primary category and value axes. */
	axisTitles: 'on' | 'off';
}

export type ChartElementKind = keyof ChartElementOptions;

const LEGEND_POSITIONS: Record<ChartElementOptions['legend'], ChartLegend['position'] | null> = {
	none: null,
	right: 'r',
	top: 't',
	left: 'l',
	bottom: 'b',
};

/** The primary category (X) and value (Y) axes of the first cartesian plot. */
function primaryAxes(model: ChartModel): { cat?: ChartAxis; val?: ChartAxis } {
	const plot = model.plots.find((p) => p.axisIds.length >= 2 && familyOfPlot(p) !== 'pie');
	if (!plot) {
		return {};
	}
	const axes = plot.axisIds.slice(0, 2).map((id) => model.axes.find((a) => a.id === id));
	const vertical = (a: ChartAxis | undefined) => a?.position === 'l' || a?.position === 'r';
	const val =
		axes.find(
			(a) =>
				a?.kind === 'val' &&
				(plot.type === 'scatter' || plot.type === 'bubble' ? vertical(a) : true),
		) ?? axes.find((a) => a?.kind === 'val');
	const cat = axes.find((a) => a && a !== val);
	return { cat, val };
}

function isPie(plot: ChartPlot): boolean {
	return familyOfPlot(plot) === 'pie';
}

/**
 * Add, remove or move a chart element (Excel's Add Chart Element menu).
 * Returns `model` unchanged when the element does not apply (e.g. axis
 * titles on a pie chart).
 */
export function setChartElement<K extends ChartElementKind>(
	model: ChartModel,
	element: K,
	option: ChartElementOptions[K],
): ChartModel {
	const next = cloneChart(model);
	switch (element) {
		case 'chartTitle':
			if (option === 'none') {
				delete next.title;
				next.autoTitleDeleted = true;
			} else {
				next.title = { ...(next.title ?? { style: labelText(14) }), overlay: option === 'overlay' };
				delete next.title.layout;
				next.autoTitleDeleted = false;
			}
			break;
		case 'legend': {
			const position = LEGEND_POSITIONS[option as ChartElementOptions['legend']];
			if (position === null) {
				delete next.legend;
			} else {
				next.legend = { ...(next.legend ?? { textStyle: labelText(9) }), position, overlay: false };
				delete next.legend.layout;
			}
			break;
		}
		case 'dataLabels':
			for (const plot of next.plots) {
				delete plot.dataLabels;
				for (const series of plot.series) {
					if (option === 'none') {
						delete series.dataLabels;
					} else {
						series.dataLabels = {
							showLegendKey: false,
							showValue: !isPie(plot),
							showCategoryName: false,
							showSeriesName: false,
							showPercent: isPie(plot),
							showBubbleSize: false,
						};
					}
				}
			}
			break;
		case 'gridlines': {
			const { val } = primaryAxes(next);
			if (!val) {
				return model;
			}
			if (option === 'on') {
				val.majorGridlines ??= hairline(GRID_COLOR);
			} else {
				delete val.majorGridlines;
			}
			break;
		}
		case 'axisTitles': {
			const axes = Object.values(primaryAxes(next)).filter((a): a is ChartAxis => Boolean(a));
			if (!axes.length) {
				return model;
			}
			for (const axis of axes) {
				if (option === 'off') {
					delete axis.title;
				} else if (!axis.title) {
					const vertical = axis.position === 'l' || axis.position === 'r';
					axis.title = { overlay: false, style: { size: 10, bold: false, color: TEXT_COLOR } };
					if (vertical) {
						axis.title.rotation = -90;
					}
				}
			}
			break;
		}
	}
	next.dirty = true;
	return next;
}

const LEGEND_OPTIONS: Record<ChartLegend['position'], ChartElementOptions['legend']> = {
	r: 'right',
	tr: 'right',
	t: 'top',
	l: 'left',
	b: 'bottom',
};

/** The current option of every Add Chart Element entry (for menu check marks). */
export function chartElementState(model: ChartModel): ChartElementOptions {
	const { cat, val } = primaryAxes(model);
	const shown = (l: ChartPlot['dataLabels']) =>
		Boolean(
			l && !l.deleted && (l.showValue || l.showPercent || l.showCategoryName || l.showSeriesName),
		);
	const labels = model.plots.some(
		(p) => shown(p.dataLabels) || p.series.some((s) => shown(s.dataLabels)),
	);
	return {
		chartTitle: model.title ? (model.title.overlay ? 'overlay' : 'above') : 'none',
		legend: model.legend ? LEGEND_OPTIONS[model.legend.position] : 'none',
		dataLabels: labels ? 'show' : 'none',
		gridlines: val?.majorGridlines ? 'on' : 'off',
		axisTitles: cat?.title || val?.title ? 'on' : 'off',
	};
}

/** Number of quick layouts `applyQuickLayout` offers. */
export const CHART_QUICK_LAYOUT_COUNT = 6;

type LayoutSpec = [
	ChartElementOptions['chartTitle'],
	ChartElementOptions['legend'],
	ChartElementOptions['dataLabels'],
	ChartElementOptions['gridlines'],
	ChartElementOptions['axisTitles'],
];

const QUICK_LAYOUTS: Record<number, LayoutSpec> = {
	1: ['above', 'right', 'none', 'on', 'off'],
	2: ['above', 'top', 'show', 'off', 'off'],
	3: ['above', 'bottom', 'none', 'on', 'off'],
	4: ['none', 'bottom', 'show', 'off', 'off'],
	5: ['above', 'none', 'none', 'on', 'on'],
	6: ['above', 'right', 'none', 'on', 'on'],
};

/** Quick Layout 1-6: a preset of title, legend, data labels, gridlines and axis titles. */
export function applyQuickLayout(model: ChartModel, layout: number): ChartModel {
	const spec = QUICK_LAYOUTS[Math.min(CHART_QUICK_LAYOUT_COUNT, Math.max(1, Math.round(layout)))];
	let next = setChartElement(model, 'chartTitle', spec[0]);
	next = setChartElement(next, 'legend', spec[1]);
	next = setChartElement(next, 'dataLabels', spec[2]);
	next = setChartElement(next, 'gridlines', spec[3]);
	next = setChartElement(next, 'axisTitles', spec[4]);
	return next === model ? cloneChart({ ...model, dirty: true }) : next;
}

const CREATABLE: Record<string, CreatableChartType> = {
	columnClustered: 'column',
	columnStacked: 'column',
	columnPercent: 'column',
	barClustered: 'bar',
	barStacked: 'bar',
	barPercent: 'bar',
	line: 'line',
	lineMarkers: 'line',
	area: 'area',
	areaStacked: 'area',
	pie: 'pie',
	doughnut: 'doughnut',
	scatter: 'scatter',
	scatterLines: 'scatter',
	radar: 'radar',
};

function union(areas: RangeAddress[]): RangeAddress {
	return {
		top: Math.min(...areas.map((a) => a.top)),
		left: Math.min(...areas.map((a) => a.left)),
		bottom: Math.max(...areas.map((a) => a.bottom)),
		right: Math.max(...areas.map((a) => a.right)),
	};
}

const cellCount = (r: RangeAddress) => (r.bottom - r.top + 1) * (r.right - r.left + 1);

/**
 * Switch Row/Column: rebuild the series from the block their references span,
 * plotted the other way round. Returns `null` when that is not possible
 * (combo charts, references to several sheets or to scattered cells,
 * unreadable references).
 */
export function switchRowColumn(
	model: ChartModel,
	workbook: Workbook,
	sheetIndex: number,
): ChartModel | null {
	const target = chartTypeTargetOf(model);
	const plot = model.plots[0];
	if (model.plots.length !== 1 || !target || !plot.series.length) {
		return null;
	}
	const refs: string[] = [];
	for (const s of plot.series) {
		if (!s.values?.ref) {
			return null;
		}
		refs.push(s.values.ref);
		for (const ref of [s.categories?.ref, s.name?.ref]) {
			if (ref) {
				refs.push(ref);
			}
		}
	}
	const areas = refs.map((ref) => resolveChartRef(ref, workbook, sheetIndex));
	if (areas.some((a) => !a || a.length !== 1)) {
		return null;
	}
	const flat = areas.map((a) => a![0]);
	const dataSheet = flat[0].sheetIndex;
	if (flat.some((a) => a.sheetIndex !== dataSheet)) {
		return null;
	}
	const block = union(flat.map((a) => a.range));
	const covered = new Set<string>();
	for (const { range } of flat) {
		if (cellCount(range) > 100_000) {
			return null;
		}
		for (let r = range.top; r <= range.bottom; r++) {
			for (let c = range.left; c <= range.right; c++) {
				covered.add(`${r}:${c}`);
			}
		}
	}
	// Every cell of the block must be part of the chart (except the empty top-left corner).
	if (covered.size < cellCount(block) - 1) {
		return null;
	}
	const values = flat[0].range;
	const inColumns =
		values.left === values.right && (values.bottom > values.top || values.top > block.top);
	const fresh = createChartFromRange(workbook, dataSheet, block, CREATABLE[target], {
		seriesIn: inColumns ? 'rows' : 'columns',
		title: false,
		legend: false,
	});
	const freshPlot = fresh.plots[0];
	if (!freshPlot?.series.length) {
		return null;
	}
	const count = freshPlot.series.length;
	const next = cloneChart(model);
	next.plots[0] = {
		...next.plots[0],
		series: freshPlot.series.map((s, i) =>
			restyleSeries(s, freshPlot, target, autoChartColor(i, count, 2)),
		),
	};
	next.dirty = true;
	return next;
}

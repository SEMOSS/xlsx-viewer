/**
 * Value transformations for plotting: clustered values, stacked running
 * totals (positive and negative bars stack separately, like Excel) and 100%
 * stacks.
 */
import type { ChartPlot } from '../../types/chart';
import type { ChartPlotData } from '../data/types';

export interface StackedValue {
	/** The point's own value (`null` gap, `NaN` skipped point). */
	raw: number | null;
	/** Start of the drawn segment in axis units (bars: base; areas: lower edge). */
	from: number;
	/** End of the drawn segment in axis units (the plotted value). */
	to: number;
	/** Share of the category total in [-1, 1] for 100% stacks. */
	share?: number;
}

export function isStacked(plot: ChartPlot): boolean {
	return plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
}

export function isPercentStacked(plot: ChartPlot): boolean {
	return plot.grouping === 'percentStacked';
}

/** Number of points (categories) in a plot. */
export function pointCount(plotData: ChartPlotData): number {
	let n = 0;
	for (const s of plotData.series) {
		if (!s.hidden) {
			n = Math.max(n, s.values.length, s.categories.length);
		}
	}
	return n;
}

function isNum(v: number | null): v is number {
	return v !== null && Number.isFinite(v);
}

/**
 * Plotted segments for every series (in `plot.series` order) and point.
 * `crossValue` is where unstacked bars start (the category axis crossing).
 */
export function stackValues(
	plot: ChartPlot,
	plotData: ChartPlotData,
	crossValue = 0,
): StackedValue[][] {
	const n = pointCount(plotData);
	const stacked = isStacked(plot) && plot.type !== 'scatter' && plot.type !== 'bubble';
	const percent = stacked && isPercentStacked(plot);
	const bars = plot.type === 'bar';
	const totals: number[] = new Array(n).fill(0);
	if (percent) {
		for (const s of plotData.series) {
			if (s.hidden) {
				continue;
			}
			s.values.forEach((v, i) => {
				if (isNum(v)) {
					totals[i] += Math.abs(v);
				}
			});
		}
	}
	const pos: number[] = new Array(n).fill(0);
	const neg: number[] = new Array(n).fill(0);
	return plotData.series.map((s) => {
		const out: StackedValue[] = [];
		for (let i = 0; i < n; i++) {
			const raw = s.hidden ? null : (s.values[i] ?? null);
			if (!stacked) {
				out.push({ raw, from: bars ? crossValue : 0, to: isNum(raw) ? raw : 0 });
				continue;
			}
			let v = isNum(raw) ? raw : 0;
			let share: number | undefined;
			if (percent) {
				share = totals[i] ? v / totals[i] : 0;
				v = share;
			}
			let from: number;
			if (bars) {
				// Bars stack positives up and negatives down from zero independently.
				from = v >= 0 ? pos[i] : neg[i];
				if (v >= 0) {
					pos[i] += v;
				} else {
					neg[i] += v;
				}
			} else {
				from = pos[i];
				pos[i] += v;
			}
			const item: StackedValue = { raw, from, to: from + v };
			if (share !== undefined) {
				item.share = share;
			}
			out.push(item);
		}
		return out;
	});
}

/** Data extent of a plot in axis units, or `null` when it has no numeric data. */
export function plotValueRange(
	plot: ChartPlot,
	plotData: ChartPlotData,
): { min: number; max: number } | null {
	let min = Infinity;
	let max = -Infinity;
	const take = (v: number) => {
		if (Number.isFinite(v)) {
			min = Math.min(min, v);
			max = Math.max(max, v);
		}
	};
	if (isStacked(plot) && plot.type !== 'scatter' && plot.type !== 'bubble') {
		for (const series of stackValues(plot, plotData)) {
			for (const p of series) {
				if (isNum(p.raw)) {
					take(p.from);
					take(p.to);
				}
			}
		}
	} else {
		for (const s of plotData.series) {
			if (s.hidden) {
				continue;
			}
			for (const v of s.values) {
				if (isNum(v)) {
					take(v);
				}
			}
		}
	}
	return min <= max ? { min, max } : null;
}

/** X extent of a scatter / bubble plot (categories that are not numbers count 1..n). */
export function plotXRange(plotData: ChartPlotData): { min: number; max: number } | null {
	let min = Infinity;
	let max = -Infinity;
	for (const s of plotData.series) {
		if (s.hidden) {
			continue;
		}
		s.values.forEach((v, i) => {
			if (!isNum(v)) {
				return;
			}
			const x = xValueAt(s, i);
			if (Number.isFinite(x)) {
				min = Math.min(min, x);
				max = Math.max(max, x);
			}
		});
	}
	return min <= max ? { min, max } : null;
}

/** X value of point `i` of a scatter series: the numeric category, else its 1-based position. */
export function xValueAt(s: ChartPlotData['series'][number], i: number): number {
	if (s.categoriesNumeric && s.xValues.length) {
		return s.xValues[i] ?? NaN;
	}
	return i + 1;
}

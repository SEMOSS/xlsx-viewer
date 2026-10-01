/**
 * Axis groups for cartesian plots: which plots share a category (X) axis and
 * a value (Y) axis, their orientation and the data they span. A combo chart
 * with a secondary axis has two groups.
 */
import type { ChartAxis, ChartPlot } from '../../types/chart';
import type { ChartSeriesData } from '../data/types';
import type { LayoutContext } from './context';
import { type DateScale, dateFraction } from './date-scale';
import { type ValueScale, scaleFraction } from './scale';
import { isPercentStacked, plotValueRange, plotXRange, pointCount } from './stacking';
import type { ChartBox } from './types';

export const CARTESIAN_TYPES = new Set([
	'bar',
	'line',
	'area',
	'scatter',
	'bubble',
	'stock',
	'surface',
]);

export interface CatAxisInfo {
	axis: ChartAxis;
	/** `value`: the X axis of scatter / bubble charts. */
	mode: 'category' | 'date' | 'value';
	count: number;
	/** Points sit between tick marks (bars) rather than on them. */
	between: boolean;
	labels: string[];
	levels?: string[][];
	/** Format for numeric categories / X values. */
	format: string;
	/** Date serials (date axes) or X values (scatter), per point. */
	positions: number[];
	dateScale?: DateScale;
	scale?: ValueScale;
	range?: { min: number; max: number } | null;
}

export interface ValAxisInfo {
	axis: ChartAxis;
	range: { min: number; max: number } | null;
	percent: boolean;
	format: string;
	/** Display-unit divisor (`c:dispUnits`), 1 when absent. */
	divisor: number;
	scale?: ValueScale;
}

export interface AxisGroup {
	plots: Array<{ plot: ChartPlot; index: number }>;
	cat: CatAxisInfo;
	val: ValAxisInfo;
	/** Categories run vertically (horizontal bar charts). */
	horizontal: boolean;
	xy: boolean;
	secondary: boolean;
}

const DISPLAY_UNITS: Record<string, number> = {
	hundreds: 1e2,
	thousands: 1e3,
	tenThousands: 1e4,
	hundredThousands: 1e5,
	millions: 1e6,
	tenMillions: 1e7,
	hundredMillions: 1e8,
	billions: 1e9,
	trillions: 1e12,
};

export function displayDivisor(axis: ChartAxis): number {
	const unit = axis.displayUnits?.unit;
	if (typeof unit === 'number') {
		return unit > 0 ? unit : 1;
	}
	return (unit && DISPLAY_UNITS[unit]) || 1;
}

function synth(
	id: number,
	kind: ChartAxis['kind'],
	position: ChartAxis['position'],
	cross: number,
): ChartAxis {
	return { id, kind, position, crossAxisId: cross };
}

function mergeRange(
	a: { min: number; max: number } | null,
	b: { min: number; max: number } | null,
): { min: number; max: number } | null {
	if (!a) {
		return b;
	}
	if (!b) {
		return a;
	}
	return { min: Math.min(a.min, b.min), max: Math.max(a.max, b.max) };
}

/** The series whose categories label the axis: the first one that has any. */
function labelSeries(ctx: LayoutContext, plots: AxisGroup['plots']): ChartSeriesData | undefined {
	for (const { index } of plots) {
		for (const s of ctx.data.plots[index]?.series ?? []) {
			if (!s.hidden && s.categories.length) {
				return s;
			}
		}
	}
	return ctx.data.plots[plots[0].index]?.series.find((s) => !s.hidden);
}

function buildCategory(ctx: LayoutContext, group: AxisGroup): void {
	const { cat } = group;
	let count = 0;
	for (const { index } of group.plots) {
		count = Math.max(count, pointCount(ctx.data.plots[index] ?? { series: [] }));
	}
	const src = labelSeries(ctx, group.plots);
	cat.count = count;
	cat.labels = Array.from({ length: count }, (_, i) => src?.categories[i] ?? String(i + 1));
	if (src?.categoryLevels?.length && !cat.axis.noMultiLevelLabels) {
		cat.levels = src.categoryLevels.map((lvl) =>
			Array.from({ length: count }, (_, i) => lvl[i] ?? ''),
		);
	}
	cat.format = src?.categoryFormat ?? 'General';
	cat.positions = Array.from({ length: count }, (_, i) => src?.xValues[i] ?? NaN);
	if (cat.mode === 'date' && !(src?.categoriesNumeric ?? false)) {
		cat.mode = 'category';
	}
}

/** Group the cartesian plots of the chart by the axes they use. */
export function buildAxisGroups(ctx: LayoutContext): AxisGroup[] {
	const byId = new Map(ctx.chart.axes.map((a) => [a.id, a] as const));
	const groups: AxisGroup[] = [];
	const keys = new Map<string, AxisGroup>();
	ctx.chart.plots.forEach((plot, index) => {
		if (!CARTESIAN_TYPES.has(plot.type)) {
			return;
		}
		const xy = plot.type === 'scatter' || plot.type === 'bubble';
		let a0 = byId.get(plot.axisIds[0]);
		let a1 = byId.get(plot.axisIds[1]);
		if (!a0 || !a1) {
			a0 = synth(-1, xy ? 'val' : 'cat', 'b', -2);
			a1 = synth(-2, 'val', 'l', -1);
		}
		if (xy ? a0.position === 'l' || a0.position === 'r' : a0.kind === 'val' && a1.kind !== 'val') {
			[a0, a1] = [a1, a0];
		}
		const key = `${a0.id}:${a1.id}`;
		let group = keys.get(key);
		if (!group) {
			group = {
				plots: [],
				cat: {
					axis: a0,
					mode: xy ? 'value' : a0.kind === 'date' ? 'date' : 'category',
					count: 0,
					between: true,
					labels: [],
					format: 'General',
					positions: [],
				},
				val: {
					axis: a1,
					range: null,
					percent: false,
					format: 'General',
					divisor: displayDivisor(a1),
				},
				horizontal: false,
				xy,
				secondary: groups.length > 0,
			};
			keys.set(key, group);
			groups.push(group);
		}
		group.plots.push({ plot, index });
		if (plot.type === 'bar' && plot.barDirection === 'bar') {
			group.horizontal = true;
		}
	});
	for (const group of groups) {
		finishGroup(ctx, group);
	}
	return groups;
}

function finishGroup(ctx: LayoutContext, group: AxisGroup): void {
	const { cat, val } = group;
	const first = ctx.data.plots[group.plots[0].index]?.series.find((s) => !s.hidden);
	if (group.xy) {
		for (const { index } of group.plots) {
			cat.range = mergeRange(
				cat.range ?? null,
				plotXRange(ctx.data.plots[index] ?? { series: [] }),
			);
		}
		cat.format = first?.categoryFormat ?? 'General';
		cat.between = false;
	} else {
		buildCategory(ctx, group);
		const crossBetween = val.axis.crossBetween;
		const hasBars = group.plots.some((p) => p.plot.type === 'bar' || p.plot.type === 'stock');
		const areaOnly = group.plots.every((p) => p.plot.type === 'area' || p.plot.type === 'surface');
		cat.between = hasBars || (crossBetween ? crossBetween === 'between' : !areaOnly);
	}
	for (const { plot, index } of group.plots) {
		val.range = mergeRange(
			val.range,
			plotValueRange(plot, ctx.data.plots[index] ?? { series: [] }),
		);
		if (isPercentStacked(plot) && plot.type !== 'scatter' && plot.type !== 'bubble') {
			val.percent = true;
		}
	}
	const numFmt = val.axis.numFmt;
	if (numFmt && !numFmt.sourceLinked && numFmt.code) {
		val.format = numFmt.code;
	} else if (val.percent) {
		val.format = '0%';
	} else {
		val.format = first?.valueFormat ?? 'General';
	}
	if (group.xy) {
		const xf = cat.axis.numFmt;
		if (xf && !xf.sourceLinked && xf.code) {
			cat.format = xf.code;
		}
	} else if (
		cat.axis.numFmt &&
		!cat.axis.numFmt.sourceLinked &&
		cat.axis.numFmt.code &&
		cat.mode === 'date'
	) {
		cat.format = cat.axis.numFmt.code;
	}
}

/** Position of point `i` along the category axis, in [0, 1]. */
export function categoryFraction(cat: CatAxisInfo, i: number): number {
	if (cat.mode === 'date' && cat.dateScale) {
		return dateFraction(cat.dateScale, cat.positions[i]);
	}
	if (cat.between) {
		return (i + 0.5) / Math.max(1, cat.count);
	}
	return cat.count > 1 ? i / (cat.count - 1) : 0.5;
}

/** Width of one category slot as a fraction of the axis (bar widths). */
export function slotFraction(cat: CatAxisInfo): number {
	if (cat.mode === 'date' && cat.dateScale) {
		return cat.dateScale.slotDays / (cat.dateScale.max - cat.dateScale.min || 1);
	}
	return 1 / Math.max(1, cat.between ? cat.count : cat.count - 1 || 1);
}

/** Map a fraction along the category axis to a pixel coordinate. */
export function catPixel(group: AxisGroup, inner: ChartBox, fraction: number): number {
	const f = group.cat.axis.orientation === 'maxMin' ? 1 - fraction : fraction;
	return group.horizontal ? inner.y + inner.height - f * inner.height : inner.x + f * inner.width;
}

/** Map a value to a pixel coordinate along the value axis. */
export function valPixel(group: AxisGroup, inner: ChartBox, value: number): number {
	const scale = group.val.scale;
	if (!scale) {
		return group.horizontal ? inner.x : inner.y + inner.height;
	}
	let f = scaleFraction(scale, value);
	if (!Number.isFinite(f)) {
		f = 0;
	}
	if (group.val.axis.orientation === 'maxMin') {
		f = 1 - f;
	}
	return group.horizontal ? inner.x + f * inner.width : inner.y + inner.height - f * inner.height;
}

/** X-axis pixel for scatter plots. */
export function xPixel(group: AxisGroup, inner: ChartBox, x: number): number {
	const scale = group.cat.scale;
	if (!scale) {
		return inner.x;
	}
	let f = scaleFraction(scale, x);
	if (group.cat.axis.orientation === 'maxMin') {
		f = 1 - f;
	}
	return inner.x + f * inner.width;
}

/** Value where the category axis crosses the value axis (bars grow from here). */
export function crossValue(group: AxisGroup): number {
	const scale = group.val.scale;
	const axis = group.cat.axis;
	if (!scale) {
		return 0;
	}
	if (axis.crossesAt !== undefined) {
		return Math.min(scale.max, Math.max(scale.min, axis.crossesAt));
	}
	if (axis.crosses === 'max') {
		return scale.max;
	}
	if (axis.crosses === 'min' || scale.logBase) {
		return scale.min;
	}
	return Math.min(scale.max, Math.max(scale.min, 0));
}

/** Fraction along the category (X) axis where the value axis crosses it. */
export function valueAxisCrossFraction(group: AxisGroup): number {
	const axis = group.val.axis;
	if (group.xy) {
		const scale = group.cat.scale;
		if (!scale) {
			return 0;
		}
		const at =
			axis.crossesAt !== undefined
				? axis.crossesAt
				: axis.crosses === 'max'
					? scale.max
					: axis.crosses === 'min' || scale.logBase
						? scale.min
						: Math.min(scale.max, Math.max(scale.min, 0));
		return scaleFraction(scale, at);
	}
	if (axis.crosses === 'max') {
		return 1;
	}
	if (axis.crossesAt !== undefined && group.cat.count > 0) {
		const n = axis.crossesAt - 1;
		return group.cat.between ? n / group.cat.count : n / Math.max(1, group.cat.count - 1);
	}
	return 0;
}

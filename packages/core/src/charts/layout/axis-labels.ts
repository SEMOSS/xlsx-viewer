/**
 * Tick label planning: value-axis scales and labels, and category labels with
 * Excel's crowding rules (wrap, then slant 45 degrees, then skip labels).
 */
import type { ChartAxis } from '../../types/chart';
import { type AxisGroup, categoryFraction, slotFraction } from './axis-model';
import {
	type LayoutContext,
	type ResolvedText,
	formatNumber,
	formatNumberColored,
	resolveText,
} from './context';
import { computeDateScale, dateFraction } from './date-scale';
import { computeValueScale, maxIntervalsFor, scaleFraction } from './scale';
import { ellipsize, wrapText } from './text';

/** Gap between the axis line and its labels, in pixels. */
export const LABEL_GAP = 5;

export interface AxisLabel {
	/** Position along the axis in [0, 1] (before orientation reversal). */
	fraction: number;
	lines: string[];
	width: number;
	/** Color from the number format (`[Red]`), when it has one. */
	color?: string;
}

export interface LevelGroup {
	start: number;
	end: number;
	text: string;
}

export interface AxisLabelPlan {
	visible: boolean;
	style: ResolvedText;
	labels: AxisLabel[];
	/** Clockwise rotation in degrees (negative slants up to the right). */
	rotation: number;
	/** Space perpendicular to the axis taken by the labels, including the gap. */
	extent: number;
	/** Outer category levels (outer first), as fraction spans. */
	levels?: LevelGroup[][];
}

/** Text style of tick labels. */
export function axisLabelStyle(ctx: LayoutContext, axis: ChartAxis): ResolvedText {
	return resolveText(ctx, { sizePt: 10 }, axis.textStyle);
}

function labelsVisible(axis: ChartAxis): boolean {
	return !axis.deleted && axis.tickLabelPosition !== 'none';
}

/** Compute the value scale for an axis of `length` pixels and plan its labels. */
export function planValueAxis(
	ctx: LayoutContext,
	group: AxisGroup,
	which: 'val' | 'x',
	length: number,
	horizontal: boolean,
): AxisLabelPlan {
	const info = which === 'val' ? group.val : group.cat;
	const axis = info.axis;
	const style = axisLabelStyle(ctx, axis);
	const range = which === 'val' ? group.val.range : (group.cat.range ?? null);
	const divisor = which === 'val' ? group.val.divisor : 1;
	const format = which === 'val' ? group.val.format : group.cat.format;
	const percent = which === 'val' && group.val.percent;
	const label = (t: number): AxisLabel => {
		const f = formatNumberColored(ctx, t / divisor, format);
		const out: AxisLabel = {
			fraction: scaleFraction(scale, t),
			lines: [f.text],
			width: ctx.measure(f.text, style.font),
		};
		if (f.color) {
			out.color = f.color;
		}
		return out;
	};
	const text = (v: number) => formatNumber(ctx, v / divisor, format);
	const scaleOpts = {
		min: axis.min,
		max: axis.max,
		majorUnit: axis.majorUnit,
		minorUnit: axis.minorUnit,
		logBase: axis.logBase,
		percent,
	};
	const lo = range?.min ?? 0;
	const hi = range?.max ?? 0;
	let spacing: number;
	if (horizontal) {
		const sample = computeValueScale(lo, hi, scaleOpts);
		const widest = Math.max(...sample.ticks.map((t) => ctx.measure(text(t), style.font)));
		spacing = widest + style.sizePx * 1.2;
	} else {
		spacing = style.lineHeight * 1.25;
	}
	let scale = computeValueScale(lo, hi, {
		...scaleOpts,
		maxIntervals: maxIntervalsFor(length, spacing),
	});
	let labels = scale.ticks.map(label);
	if (horizontal && axis.majorUnit === undefined && labels.length > 1) {
		const widest = Math.max(...labels.map((l) => l.width));
		const per = length / (labels.length - 1);
		if (widest + style.sizePx * 0.6 > per) {
			scale = computeValueScale(lo, hi, {
				...scaleOpts,
				maxIntervals: maxIntervalsFor(length, widest + style.sizePx * 1.2),
			});
			labels = scale.ticks.map(label);
		}
	}
	if (which === 'val') {
		group.val.scale = scale;
	} else {
		group.cat.scale = scale;
	}
	const visible = labelsVisible(axis);
	const widest = labels.reduce((m, l) => Math.max(m, l.width), 0);
	const extent = !visible ? 0 : horizontal ? style.lineHeight + LABEL_GAP : widest + LABEL_GAP + 2;
	return { visible, style, labels, rotation: 0, extent };
}

/** Group consecutive points of an outer level into labelled spans. */
function levelGroups(level: string[], fractionOf: (i: number) => [number, number]): LevelGroup[] {
	const out: LevelGroup[] = [];
	for (let i = 0; i < level.length; i++) {
		const [a, b] = fractionOf(i);
		if (level[i] !== '' || out.length === 0) {
			out.push({ start: a, end: b, text: level[i] });
		} else {
			out[out.length - 1].end = b;
		}
	}
	return out;
}

/**
 * Plan category (or date) labels for an axis of `length` pixels. `maxExtent`
 * caps the space labels may take perpendicular to the axis.
 */
export function planCategoryAxis(
	ctx: LayoutContext,
	group: AxisGroup,
	length: number,
	maxExtent: number,
): AxisLabelPlan {
	const cat = group.cat;
	const axis = cat.axis;
	const style = axisLabelStyle(ctx, axis);
	const visible = labelsVisible(axis);
	let items: Array<{ fraction: number; text: string }>;
	let slotPx: number;
	if (cat.mode === 'date') {
		const spacing = style.sizePx * 0.6 * 10;
		cat.dateScale = computeDateScale(cat.positions, {
			between: cat.between,
			baseUnit: axis.baseTimeUnit,
			majorUnit: axis.majorUnit,
			majorTimeUnit: axis.majorTimeUnit,
			min: axis.min,
			max: axis.max,
			maxIntervals: maxIntervalsFor(length, spacing),
			date1904: ctx.date1904,
		});
		const scale = cat.dateScale;
		items = scale.ticks.map((t) => ({
			fraction: dateFraction(scale, t),
			text: formatNumber(ctx, t, cat.format),
		}));
		slotPx = items.length > 1 ? Math.abs(items[1].fraction - items[0].fraction) * length : length;
	} else {
		items = cat.labels.map((text, i) => ({ fraction: categoryFraction(cat, i), text }));
		slotPx = slotFraction(cat) * length;
	}
	const lh = style.lineHeight;
	const font = style.font;
	const measure = ctx.measure;
	const widths = items.map((it) => measure(it.text, font));
	const widest = widths.reduce((m, w) => Math.max(m, w), 0);
	let rotation = axis.labelRotation ?? 0;
	let skip = axis.tickLabelSkip ?? 1;
	let lines: string[][] = items.map((it) => [it.text]);
	if (group.horizontal) {
		// Vertical category axis (bar charts): horizontal labels, wrapped to the cap, skipped when crowded.
		const cap = Math.max(20, maxExtent - LABEL_GAP);
		const maxLines = Math.max(1, Math.floor(slotPx / lh));
		lines = items.map((it) => wrapText(it.text, cap, font, measure, Math.min(2, maxLines)));
		if (axis.tickLabelSkip === undefined && slotPx > 0 && slotPx < lh) {
			skip = Math.ceil((lh * 1.1) / slotPx);
		}
		rotation = 0;
	} else if (axis.labelRotation === undefined) {
		if (widest > slotPx - 4) {
			const wrapped = items.map((it) =>
				wrapText(it.text, Math.max(1, slotPx - 4), font, measure, 3),
			);
			const wordsFit = items.every((it) =>
				it.text.split(/\s+/).every((w) => measure(w, font) <= slotPx - 4),
			);
			if (wordsFit && slotPx >= lh) {
				lines = wrapped;
			} else {
				rotation = -45;
				const needed = lh * Math.SQRT2;
				if (axis.tickLabelSkip === undefined && slotPx < needed) {
					skip = Math.ceil(needed / Math.max(1, slotPx));
				}
			}
		}
	} else if (axis.tickLabelSkip === undefined && rotation !== 0) {
		const rad = (Math.abs(rotation) * Math.PI) / 180;
		const needed = lh / Math.max(0.2, Math.sin(rad));
		if (slotPx < needed) {
			skip = Math.ceil(needed / Math.max(1, slotPx));
		}
	}
	const labels: AxisLabelPlan['labels'] = [];
	let extent = 0;
	const rad = (Math.abs(rotation) * Math.PI) / 180;
	const maxRotatedWidth =
		rotation === 0
			? Infinity
			: Math.max(10, (maxExtent - LABEL_GAP - lh * Math.cos(rad)) / Math.max(0.2, Math.sin(rad)));
	items.forEach((it, i) => {
		if (i % skip !== 0) {
			return;
		}
		let ls = lines[i];
		if (rotation !== 0) {
			ls = [ellipsize(it.text, maxRotatedWidth, font, measure)];
		}
		const w = ls.reduce((m, l) => Math.max(m, measure(l, font)), 0);
		labels.push({ fraction: it.fraction, lines: ls, width: w });
		const across = group.horizontal
			? w
			: rotation === 0
				? ls.length * lh
				: w * Math.sin(rad) + lh * Math.cos(rad);
		extent = Math.max(extent, across);
	});
	const plan: AxisLabelPlan = {
		visible,
		style,
		labels,
		rotation,
		extent: visible ? Math.min(extent, maxExtent) + LABEL_GAP : 0,
	};
	if (visible && cat.levels?.length && cat.mode === 'category') {
		const half = slotFraction(cat) / 2;
		plan.levels = cat.levels.map((lvl) =>
			levelGroups(lvl, (i) => {
				const f = categoryFraction(cat, i);
				return cat.between ? [f - half, f + half] : [f, f];
			}),
		);
		const levelExtent = group.horizontal
			? Math.max(...plan.levels.flat().map((g) => measure(g.text, font))) + LABEL_GAP * 2
			: lh + LABEL_GAP;
		plan.extent += levelExtent * plan.levels.length;
	}
	return plan;
}

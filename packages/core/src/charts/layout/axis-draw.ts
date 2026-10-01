/** Drawing of axis lines, tick marks, gridlines, tick labels and axis titles. */
import type { ChartAxis } from '../../types/chart';
import { LABEL_GAP, type AxisLabelPlan } from './axis-labels';
import {
	type AxisGroup,
	categoryFraction,
	catPixel,
	crossValue,
	slotFraction,
	valPixel,
	valueAxisCrossFraction,
} from './axis-model';
import {
	DEFAULT_AXIS_COLOR,
	DEFAULT_GRIDLINE_COLOR,
	DEFAULT_MINOR_GRIDLINE_COLOR,
	type LayoutContext,
	type ResolvedText,
	resolveText,
	strokePaint,
} from './context';
import { dateFraction } from './date-scale';
import { scaleFraction } from './scale';
import { blockHeight, textBlock } from './text-block';
import { wrapText } from './text';
import type { ChartBox, ChartLineNode, ChartRenderNode } from './types';

type Side = ChartAxis['position'];
const TICK = 5;

/** Which axis of a group: the category / X axis or the value / Y axis. */
export type AxisWhich = 'cat' | 'val';

export function axisOf(group: AxisGroup, which: AxisWhich): ChartAxis {
	return which === 'cat' ? group.cat.axis : group.val.axis;
}

/** Whether the axis runs horizontally. */
export function isHorizontalAxis(group: AxisGroup, which: AxisWhich): boolean {
	return which === 'cat' ? !group.horizontal : group.horizontal;
}

/** Side of the plot the axis labels sit on, consistent with the axis orientation. */
export function axisSide(group: AxisGroup, which: AxisWhich): Side {
	const pos = axisOf(group, which).position;
	if (isHorizontalAxis(group, which)) {
		return pos === 't' ? 't' : 'b';
	}
	return pos === 'r' ? 'r' : 'l';
}

/** Pixel along the axis for a fraction in [0, 1] (respecting reversal). */
function alongPixel(group: AxisGroup, which: AxisWhich, inner: ChartBox, fraction: number): number {
	const axis = axisOf(group, which);
	const f = axis.orientation === 'maxMin' ? 1 - fraction : fraction;
	return isHorizontalAxis(group, which)
		? inner.x + f * inner.width
		: inner.y + inner.height - f * inner.height;
}

/** Perpendicular coordinate of the axis line. */
function linePosition(group: AxisGroup, which: AxisWhich, inner: ChartBox): number {
	if (which === 'cat') {
		return valPixel(group, inner, crossValue(group));
	}
	return catPixel(group, inner, valueAxisCrossFraction(group));
}

/** Perpendicular coordinate the labels attach to (`nextTo`: the axis line; `low` / `high`: plot edges). */
function labelBase(group: AxisGroup, which: AxisWhich, inner: ChartBox): number {
	const axis = axisOf(group, which);
	const horizontal = isHorizontalAxis(group, which);
	if (axis.tickLabelPosition === 'low' || axis.tickLabelPosition === 'high') {
		const other = which === 'cat' ? group.val.axis : group.cat.axis;
		const atMin = (axis.tickLabelPosition === 'low') !== (other.orientation === 'maxMin');
		if (horizontal) {
			return atMin ? inner.y + inner.height : inner.y;
		}
		return atMin ? inner.x : inner.x + inner.width;
	}
	return linePosition(group, which, inner);
}

function outward(side: Side): 1 | -1 {
	return side === 'b' || side === 'r' ? 1 : -1;
}

function lineNode(
	horizontal: boolean,
	along: number,
	a: number,
	b: number,
): Omit<ChartLineNode, 'stroke'> {
	return horizontal
		? { type: 'line', x1: along, y1: a, x2: along, y2: b }
		: { type: 'line', x1: a, y1: along, x2: b, y2: along };
}

/** Fractions of the major ticks / gridlines along an axis. */
function tickFractions(group: AxisGroup, which: AxisWhich, minor = false): number[] {
	if (which === 'val' || group.xy) {
		const scale = which === 'val' ? group.val.scale : group.cat.scale;
		if (!scale) {
			return [];
		}
		return (minor ? scale.minorTicks : scale.ticks).map((t) => scaleFraction(scale, t));
	}
	const cat = group.cat;
	if (minor) {
		return [];
	}
	if (cat.mode === 'date' && cat.dateScale) {
		const ds = cat.dateScale;
		return ds.ticks.map((t) => dateFraction(ds, t));
	}
	if (cat.between) {
		const slot = slotFraction(cat);
		return Array.from({ length: cat.count + 1 }, (_, k) => k * slot);
	}
	return Array.from({ length: cat.count }, (_, i) => categoryFraction(cat, i));
}

/** Gridlines of both axes of a group (minor first, then major). */
export function drawGridlines(
	ctx: LayoutContext,
	group: AxisGroup,
	inner: ChartBox,
): ChartRenderNode[] {
	const out: ChartRenderNode[] = [];
	for (const minor of [true, false]) {
		for (const which of ['val', 'cat'] as const) {
			const axis = axisOf(group, which);
			const spec = minor ? axis.minorGridlines : axis.majorGridlines;
			if (!spec) {
				continue;
			}
			const paint = strokePaint(ctx, spec === true ? undefined : spec, {
				color: minor ? DEFAULT_MINOR_GRIDLINE_COLOR : DEFAULT_GRIDLINE_COLOR,
				width: 1,
			});
			if (!paint.stroke) {
				continue;
			}
			const horizontal = isHorizontalAxis(group, which);
			for (const f of tickFractions(group, which, minor)) {
				if (f < -1e-6 || f > 1 + 1e-6) {
					continue;
				}
				const p = alongPixel(group, which, inner, f);
				const node = horizontal
					? lineNode(true, p, inner.y, inner.y + inner.height)
					: lineNode(false, p, inner.x, inner.x + inner.width);
				out.push({ ...node, ...paint, role: 'gridline' } as ChartLineNode);
			}
		}
	}
	return out;
}

/** Axis line and tick marks. */
export function drawAxisLine(
	ctx: LayoutContext,
	group: AxisGroup,
	which: AxisWhich,
	inner: ChartBox,
): ChartRenderNode[] {
	const axis = axisOf(group, which);
	if (axis.deleted) {
		return [];
	}
	const paint = strokePaint(ctx, axis.line, { color: DEFAULT_AXIS_COLOR, width: 1 });
	if (!paint.stroke) {
		return [];
	}
	const horizontal = isHorizontalAxis(group, which);
	const pos = linePosition(group, which, inner);
	const out: ChartRenderNode[] = [];
	const [a, b] = horizontal ? [inner.x, inner.x + inner.width] : [inner.y, inner.y + inner.height];
	out.push(
		horizontal
			? { type: 'line', x1: a, y1: pos, x2: b, y2: pos, ...paint, role: 'axis-line' }
			: { type: 'line', x1: pos, y1: a, x2: pos, y2: b, ...paint, role: 'axis-line' },
	);
	const dir = outward(axisSide(group, which));
	for (const minor of [false, true]) {
		const mark = (minor ? axis.minorTickMark : axis.majorTickMark) ?? (minor ? 'none' : 'out');
		if (mark === 'none') {
			continue;
		}
		const len = minor ? TICK - 2 : TICK;
		const from = mark === 'in' ? -len * dir : mark === 'cross' ? -len * dir : 0;
		const to = mark === 'in' ? 0 : len * dir;
		for (const f of tickFractions(group, which, minor)) {
			if (f < -1e-6 || f > 1 + 1e-6) {
				continue;
			}
			const along = alongPixel(group, which, inner, f);
			const node = horizontal
				? lineNode(true, along, pos + from, pos + to)
				: lineNode(false, along, pos + from, pos + to);
			out.push({ ...node, ...paint, dash: undefined, role: 'tick' } as ChartLineNode);
		}
	}
	return out;
}

/** Tick labels (and multi-level category groups). */
export function drawAxisLabels(
	group: AxisGroup,
	which: AxisWhich,
	inner: ChartBox,
	plan: AxisLabelPlan,
): ChartRenderNode[] {
	if (!plan.visible) {
		return [];
	}
	const horizontal = isHorizontalAxis(group, which);
	const side = axisSide(group, which);
	const dir = outward(side);
	const base = labelBase(group, which, inner);
	const style = plan.style;
	const out: ChartRenderNode[] = [];
	const meta = { role: 'tick-label' as const };
	let maxAcross = 0;
	for (const label of plan.labels) {
		if (label.fraction < -1e-6 || label.fraction > 1 + 1e-6) {
			continue;
		}
		const first = out.length;
		const along = alongPixel(group, which, inner, label.fraction);
		const h = blockHeight(style, label.lines.length);
		if (horizontal) {
			if (plan.rotation === 0) {
				const top = dir > 0 ? base + LABEL_GAP : base - LABEL_GAP - h;
				out.push(textBlock(style, label.lines, along, top, 'middle', meta));
				maxAcross = Math.max(maxAcross, h);
			} else {
				const rad = (Math.abs(plan.rotation) * Math.PI) / 180;
				const pivotY = base + dir * (LABEL_GAP + (style.lineHeight / 2) * Math.cos(rad));
				const anchor = plan.rotation < 0 === dir > 0 ? 'end' : 'start';
				out.push(
					textBlock(
						style,
						label.lines.slice(0, 1),
						along,
						pivotY - style.lineHeight / 2,
						anchor,
						meta,
						{
							rotation: plan.rotation,
							pivotX: along,
							pivotY,
						},
					),
				);
				maxAcross = Math.max(
					maxAcross,
					label.width * Math.sin(rad) + style.lineHeight * Math.cos(rad),
				);
			}
		} else {
			const x = base + dir * LABEL_GAP;
			out.push(textBlock(style, label.lines, x, along - h / 2, dir < 0 ? 'end' : 'start', meta));
			maxAcross = Math.max(maxAcross, label.width);
		}
		const node = out[first];
		if (label.color && node?.type === 'text') {
			node.fill = label.color;
		}
	}
	if (plan.levels) {
		let offset = LABEL_GAP + maxAcross + LABEL_GAP;
		for (const level of [...plan.levels].reverse()) {
			let levelAcross = 0;
			for (const g of level) {
				const p0 = alongPixel(group, which, inner, g.start);
				const p1 = alongPixel(group, which, inner, g.end);
				const mid = (p0 + p1) / 2;
				if (horizontal) {
					const top = dir > 0 ? base + offset : base - offset - style.lineHeight;
					out.push(textBlock(style, [g.text], mid, top, 'middle', meta));
					levelAcross = Math.max(levelAcross, style.lineHeight);
				} else {
					const x = base + dir * offset;
					out.push(
						textBlock(
							style,
							[g.text],
							x,
							mid - style.lineHeight / 2,
							dir < 0 ? 'end' : 'start',
							meta,
						),
					);
				}
				for (const edge of [p0, p1]) {
					const a = base + dir * (offset - LABEL_GAP);
					const b = base + dir * (offset + style.lineHeight);
					const node = horizontal ? lineNode(true, edge, a, b) : lineNode(false, edge, a, b);
					out.push({
						...node,
						stroke: DEFAULT_AXIS_COLOR,
						strokeWidth: 1,
						role: 'tick',
					} as ChartLineNode);
				}
			}
			offset += (horizontal ? levelAcross : style.lineHeight) + LABEL_GAP;
		}
	}
	return out;
}

export interface AxisTitlePlan {
	lines: string[];
	style: ResolvedText;
	rotation: number;
	/** Widest line in pixels. */
	width: number;
	/** Space taken perpendicular to the axis, including the gap. */
	extent: number;
}

/** Plan an axis title (text, rotation and the space it needs). */
export function planAxisTitle(
	ctx: LayoutContext,
	axis: ChartAxis,
	horizontalAxis: boolean,
	maxLength: number,
): AxisTitlePlan | undefined {
	if (!axis.title || axis.deleted) {
		return undefined;
	}
	const text = ctx.data.titles?.axes?.[axis.id] ?? axis.title.text ?? 'Axis Title';
	if (!text.trim()) {
		return undefined;
	}
	const style = resolveText(ctx, { sizePt: 10 }, axis.title.style);
	const rotation = axis.title.rotation ?? (horizontalAxis ? 0 : -90);
	const lines = wrapText(text, Math.max(20, maxLength), style.font, ctx.measure, 3);
	const width = lines.reduce((m, l) => Math.max(m, ctx.measure(l, style.font)), 0);
	// A vertical axis with an unrotated title needs the text width beside the axis.
	const across = !horizontalAxis && rotation === 0 ? width : blockHeight(style, lines.length);
	return { lines, style, rotation, width, extent: across + LABEL_GAP };
}

/** Draw an axis title centred along `inner` at distance `edge` (the outer edge of the side). */
export function drawAxisTitle(
	plan: AxisTitlePlan,
	side: Side,
	inner: ChartBox,
	edge: number,
): ChartRenderNode {
	const h = blockHeight(plan.style, plan.lines.length);
	const meta = { role: 'axis-title' as const };
	if (side === 'b' || side === 't') {
		const top = side === 'b' ? edge - h : edge;
		return textBlock(plan.style, plan.lines, inner.x + inner.width / 2, top, 'middle', meta);
	}
	const cy = inner.y + inner.height / 2;
	if (plan.rotation === 0) {
		return textBlock(
			plan.style,
			plan.lines,
			edge,
			cy - h / 2,
			side === 'l' ? 'start' : 'end',
			meta,
		);
	}
	const cx = side === 'l' ? edge + h / 2 : edge - h / 2;
	return textBlock(plan.style, plan.lines, cx, cy - h / 2, 'middle', meta, {
		rotation: plan.rotation,
		pivotX: cx,
		pivotY: cy,
	});
}

const UNIT_LABELS: Record<string, string> = {
	hundreds: 'Hundreds',
	thousands: 'Thousands',
	tenThousands: 'x 10000',
	hundredThousands: 'x 100000',
	millions: 'Millions',
	tenMillions: 'x 10000000',
	hundredMillions: 'x 100000000',
	billions: 'Billions',
	trillions: 'Trillions',
};

/** The display-units label (`Thousands`, `Millions`, ...) of a value axis, when shown. */
export function planUnitsLabel(ctx: LayoutContext, group: AxisGroup): AxisTitlePlan | undefined {
	const axis = group.val.axis;
	const units = axis.displayUnits;
	if (!units?.showLabel || axis.deleted || axis.tickLabelPosition === 'none') {
		return undefined;
	}
	const text =
		typeof units.unit === 'number' ? `x ${units.unit}` : (UNIT_LABELS[units.unit] ?? units.unit);
	const style = resolveText(ctx, { sizePt: 10 }, axis.textStyle);
	const width = ctx.measure(text, style.font);
	return { lines: [text], style, rotation: -90, width, extent: style.lineHeight + LABEL_GAP };
}

/** Draw the display-units label beside the tick labels, at the far end of the value axis. */
export function drawUnitsLabel(
	plan: AxisTitlePlan,
	group: AxisGroup,
	inner: ChartBox,
	labelExtent: number,
): ChartRenderNode {
	const side = axisSide(group, 'val');
	const meta = { role: 'axis-title' as const };
	if (side === 'b' || side === 't') {
		const top =
			side === 'b' ? inner.y + inner.height + labelExtent : inner.y - labelExtent - plan.extent;
		return textBlock(plan.style, plan.lines, inner.x + inner.width, top, 'end', meta);
	}
	const half = plan.style.lineHeight / 2;
	const cx =
		side === 'l' ? inner.x - labelExtent - half : inner.x + inner.width + labelExtent + half;
	return textBlock(plan.style, plan.lines, cx, inner.y - half, 'end', meta, {
		rotation: -90,
		pivotX: cx,
		pivotY: inner.y,
	});
}

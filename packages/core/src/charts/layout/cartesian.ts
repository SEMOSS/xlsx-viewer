/**
 * Cartesian charts (bar, line, area, scatter, bubble, stock, surface): solve
 * the inner plot rectangle around axis labels and titles, then draw gridlines,
 * series, axes and labels in Excel's stacking order.
 */
import type { ChartPlot } from '../../types/chart';
import { type AxisLabelPlan, planCategoryAxis, planValueAxis } from './axis-labels';
import {
	type AxisTitlePlan,
	axisSide,
	drawAxisLabels,
	drawAxisLine,
	drawAxisTitle,
	drawGridlines,
	drawUnitsLabel,
	isHorizontalAxis,
	planAxisTitle,
	planUnitsLabel,
} from './axis-draw';
import { type AxisGroup, buildAxisGroups } from './axis-model';
import { type LayoutContext, fillPaint, strokePaint } from './context';
import { drawBarPlot } from './plot-bar';
import { drawAreaPlot, drawLinePlot } from './plot-line';
import { drawBubblePlot, drawScatterPlot } from './plot-scatter';
import type { ChartBox, ChartRenderNode } from './types';

type Side = 'l' | 'r' | 't' | 'b';

interface GroupPlan {
	group: AxisGroup;
	cat: AxisLabelPlan;
	val: AxisLabelPlan;
	catTitle?: AxisTitlePlan;
	valTitle?: AxisTitlePlan;
	units?: AxisTitlePlan;
}

const Z_ORDER: Record<string, number> = {
	area: 0,
	surface: 0,
	bar: 1,
	stock: 2,
	line: 2,
	scatter: 3,
	bubble: 3,
};

function shrink(box: ChartBox, m: Record<Side, number>): ChartBox {
	const width = Math.max(10, box.width - m.l - m.r);
	const height = Math.max(10, box.height - m.t - m.b);
	return {
		x: box.x + Math.min(m.l, box.width - 10),
		y: box.y + Math.min(m.t, box.height - 10),
		width,
		height,
	};
}

function planGroups(
	ctx: LayoutContext,
	groups: AxisGroup[],
	inner: ChartBox,
	outer: ChartBox,
): GroupPlan[] {
	return groups.map((group) => {
		const catLen = group.horizontal ? inner.height : inner.width;
		const valLen = group.horizontal ? inner.width : inner.height;
		const maxCatExtent = group.horizontal ? outer.width * 0.4 : outer.height * 0.45;
		const val = planValueAxis(ctx, group, 'val', valLen, group.horizontal);
		const cat = group.xy
			? planValueAxis(ctx, group, 'x', catLen, true)
			: planCategoryAxis(ctx, group, catLen, maxCatExtent);
		return {
			group,
			cat,
			val,
			catTitle: planAxisTitle(ctx, group.cat.axis, isHorizontalAxis(group, 'cat'), catLen),
			valTitle: planAxisTitle(ctx, group.val.axis, isHorizontalAxis(group, 'val'), valLen),
			units: planUnitsLabel(ctx, group),
		};
	});
}

function marginsFor(plans: GroupPlan[]): Record<Side, number> {
	const m: Record<Side, number> = { l: 0, r: 0, t: 0, b: 0 };
	for (const p of plans) {
		const catSide = axisSide(p.group, 'cat');
		const valSide = axisSide(p.group, 'val');
		m[catSide] += p.cat.extent + (p.catTitle?.extent ?? 0);
		m[valSide] += p.val.extent + (p.valTitle?.extent ?? 0) + (p.units?.extent ?? 0);
		// Labels centred on the ends of an axis overhang the plot by half their size.
		const valHalf = p.val.visible
			? isHorizontalAxis(p.group, 'val')
				? halfWidth(p.val)
				: p.val.style.lineHeight / 2
			: 0;
		const catHalf = p.cat.visible && p.group.xy ? halfWidth(p.cat) : 0;
		if (isHorizontalAxis(p.group, 'val')) {
			m.l = Math.max(m.l, valHalf);
			m.r = Math.max(m.r, valHalf);
		} else {
			m.t = Math.max(m.t, valHalf);
			m.b = Math.max(m.b, valHalf);
		}
		m.l = Math.max(m.l, catHalf);
		m.r = Math.max(m.r, catHalf);
	}
	return m;
}

function halfWidth(plan: AxisLabelPlan): number {
	const first = plan.labels[0]?.width ?? 0;
	const last = plan.labels[plan.labels.length - 1]?.width ?? 0;
	return Math.max(first, last) / 2;
}

function drawPlot(
	ctx: LayoutContext,
	group: AxisGroup,
	inner: ChartBox,
	plot: ChartPlot,
	index: number,
	labels: ChartRenderNode[],
): ChartRenderNode[] {
	switch (plot.type) {
		case 'bar':
			return drawBarPlot(ctx, group, inner, plot, index, labels);
		case 'area':
			return drawAreaPlot(ctx, group, inner, plot, index, labels);
		case 'scatter':
			return drawScatterPlot(ctx, group, inner, plot, index, labels);
		case 'bubble':
			return drawBubblePlot(ctx, group, inner, plot, index, labels);
		default:
			return drawLinePlot(ctx, group, inner, plot, index, labels);
	}
}

export interface CartesianLayout {
	nodes: ChartRenderNode[];
	inner: ChartBox;
}

/**
 * Lay out every cartesian plot within `outer` (the plot area including axis
 * labels). With `fixedInner` (a manual inner layout) the inner rectangle is
 * taken as given, shrunk only as far as needed to keep labels inside `outer`.
 */
export function layoutCartesian(
	ctx: LayoutContext,
	outer: ChartBox,
	fixedInner?: ChartBox,
): CartesianLayout | undefined {
	const groups = buildAxisGroups(ctx);
	if (!groups.length) {
		return undefined;
	}
	let inner = fixedInner ?? shrink(outer, { l: 40, r: 10, t: 10, b: 24 });
	let plans = planGroups(ctx, groups, inner, outer);
	if (!fixedInner) {
		for (let pass = 0; pass < 2; pass++) {
			inner = shrink(outer, marginsFor(plans));
			plans = planGroups(ctx, groups, inner, outer);
		}
	} else {
		// Like Excel, keep a manually placed plot area's labels inside the chart.
		const m = marginsFor(plans);
		const x0 = Math.max(inner.x, outer.x + m.l);
		const y0 = Math.max(inner.y, outer.y + m.t);
		const x1 = Math.min(inner.x + inner.width, outer.x + outer.width - m.r);
		const y1 = Math.min(inner.y + inner.height, outer.y + outer.height - m.b);
		if (
			x1 - x0 >= 10 &&
			y1 - y0 >= 10 &&
			(x0 !== inner.x ||
				y0 !== inner.y ||
				x1 !== inner.x + inner.width ||
				y1 !== inner.y + inner.height)
		) {
			inner = { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
			plans = planGroups(ctx, groups, inner, outer);
		}
	}
	const nodes: ChartRenderNode[] = [];
	const plotPaint = {
		...fillPaint(ctx, ctx.chart.plotFill, null),
		...strokePaint(ctx, ctx.chart.plotLine, null),
	};
	if (plotPaint.fill || plotPaint.gradient || plotPaint.stroke) {
		nodes.push({ type: 'rect', ...inner, ...plotPaint, role: 'plot-area' });
	}
	for (const p of plans) {
		nodes.push(...drawGridlines(ctx, p.group, inner));
	}
	const labels: ChartRenderNode[] = [];
	for (const p of plans) {
		const sorted = [...p.group.plots].sort(
			(a, b) => (Z_ORDER[a.plot.type] ?? 2) - (Z_ORDER[b.plot.type] ?? 2),
		);
		for (const { plot, index } of sorted) {
			const filled = plot.type === 'bar' || plot.type === 'area';
			const pad = filled ? 0 : 8;
			nodes.push({
				type: 'group',
				clip: {
					x: inner.x - pad,
					y: inner.y - pad,
					width: inner.width + 2 * pad,
					height: inner.height + 2 * pad,
				},
				children: drawPlot(ctx, p.group, inner, plot, index, labels),
				plot: index,
			});
		}
	}
	const edges: Record<Side, number> = {
		l: outer.x,
		r: outer.x + outer.width,
		t: outer.y,
		b: outer.y + outer.height,
	};
	for (const p of plans) {
		nodes.push(
			...drawAxisLine(ctx, p.group, 'val', inner),
			...drawAxisLine(ctx, p.group, 'cat', inner),
		);
	}
	for (const p of plans) {
		if (p.units) {
			nodes.push(drawUnitsLabel(p.units, p.group, inner, p.val.extent));
		}
		nodes.push(
			...drawAxisLabels(p.group, 'val', inner, p.val),
			...drawAxisLabels(p.group, 'cat', inner, p.cat),
		);
		for (const [title, which] of [
			[p.catTitle, 'cat'],
			[p.valTitle, 'val'],
		] as const) {
			if (!title) {
				continue;
			}
			const side = axisSide(p.group, which);
			nodes.push(drawAxisTitle(title, side, inner, edges[side]));
			edges[side] += side === 'l' || side === 't' ? title.extent : -title.extent;
		}
	}
	nodes.push(...labels);
	return { nodes, inner };
}

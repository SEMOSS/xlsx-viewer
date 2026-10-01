/** Pie, doughnut and (approximated) pie-of-pie plots. */
import type { ChartDataLabels, ChartPlot } from '../../types/chart';
import type { LayoutContext } from './context';
import { effectiveLabels, labelSize, labelText, labelTextStyle, placeLabel } from './data-labels';
import { polar, sectorPath } from './geometry';
import { findPoint, orderedSeries, shapePaint } from './series-style';
import type { ChartBox, ChartRenderNode, ChartTextNode } from './types';

const GAP = 6;

export interface PieSlice {
	point: number;
	value: number;
	share: number;
	/** Degrees clockwise from 12 o'clock. */
	start: number;
	end: number;
	explosion: number;
}

/** Slice angles for the values of one series, starting at `firstSliceAngle`. */
export function pieSlices(
	values: ReadonlyArray<number | null>,
	firstSliceAngle = 0,
	explosionOf: (i: number) => number = () => 0,
): PieSlice[] {
	const abs = values.map((v) => (v !== null && Number.isFinite(v) ? Math.abs(v) : 0));
	const total = abs.reduce((s, v) => s + v, 0);
	const slices: PieSlice[] = [];
	if (total <= 0) {
		return slices;
	}
	let angle = ((firstSliceAngle % 360) + 360) % 360;
	abs.forEach((v, i) => {
		if (v <= 0) {
			return;
		}
		const sweep = (v / total) * 360;
		slices.push({
			point: i,
			value: values[i] as number,
			share: v / total,
			start: angle,
			end: angle + sweep,
			explosion: explosionOf(i),
		});
		angle += sweep;
	});
	return slices;
}

/** Best fit: a label goes inside its slice when it fits the chord at 62% of the radius. */
function fitsInside(
	size: { width: number; height: number },
	slice: PieSlice,
	radius: number,
): boolean {
	const r = radius * 0.62;
	const sweep = ((slice.end - slice.start) * Math.PI) / 180;
	const chord = 2 * r * Math.sin(Math.min(Math.PI / 2, sweep / 2));
	return size.width <= chord * 0.95 && size.height <= radius * 0.45 && sweep > 0.35;
}

/** Put each part of a multi-part label (category, value, percentage) on its own line. */
function wrapAtSeparator(text: string, lbl: ChartDataLabels): string {
	const sep = lbl.separator ?? ', ';
	return sep && sep !== '\n' ? text.split(sep).join('\n') : text;
}

function labelAnchor(angle: number): ChartTextNode['anchor'] {
	const a = ((angle % 360) + 360) % 360;
	if (a > 5 && a < 175) {
		return 'start';
	}
	if (a > 185 && a < 355) {
		return 'end';
	}
	return 'middle';
}

/** Draw the pie-family plot into `box`; labels are appended to `labels`. */
export function drawPiePlot(
	ctx: LayoutContext,
	box: ChartBox,
	plot: ChartPlot,
	plotIndex: number,
	labels: ChartRenderNode[],
): ChartRenderNode[] {
	const plotData = ctx.data.plots[plotIndex];
	if (!plotData) {
		return [];
	}
	const doughnut = plot.type === 'doughnut';
	const visible = orderedSeries(plot).filter(({ index }) => !plotData.series[index]?.hidden);
	// A pie shows its first series only; a doughnut shows every series as a ring.
	const rings = doughnut ? visible : visible.slice(0, 1);
	if (!rings.length) {
		return [];
	}
	const prepared = rings.map(({ series, index }) => {
		const data = plotData.series[index];
		const explosionOf = (i: number) =>
			Math.max(0, findPoint(series, i)?.explosion ?? series.explosion ?? 0);
		return {
			series,
			index,
			data,
			slices: pieSlices(data.values, plot.firstSliceAngle ?? 0, explosionOf),
		};
	});
	const cx = box.x + box.width / 2;
	const cy = box.y + box.height / 2;
	const full = Math.max(4, Math.min(box.width, box.height) / 2 - 2);
	const maxExplosion = doughnut
		? 0
		: Math.max(0, ...prepared.flatMap((r) => r.slices.map((s) => s.explosion)));
	const radiusFor = (room: number) =>
		Math.max(full * 0.45, Math.min(full, room)) / (1 + Math.min(400, maxExplosion) / 100);
	let radius = radiusFor(full);
	// Shrink the pie to make room for labels that cannot sit inside their slice.
	let outsideW = 0;
	let outsideH = 0;
	if (!doughnut) {
		for (const r of prepared) {
			for (const s of r.slices) {
				const lbl = effectiveLabels(plot, r.series, s.point);
				const position = lbl?.position ?? 'bestFit';
				if (!lbl || position === 'ctr' || position === 'inEnd') {
					continue;
				}
				const text = labelText(ctx, lbl, {
					value: s.value,
					category: r.data.categories[s.point] ?? '',
					seriesName: r.data.name,
					percent: s.share,
					valueFormat: r.data.valueFormat,
				});
				const style = labelTextStyle(ctx, plot, lbl);
				const size = labelSize(ctx, style, text);
				const wrapped = labelSize(ctx, style, wrapAtSeparator(text, lbl));
				if (
					position === 'outEnd' ||
					(!fitsInside(size, s, radius) && !fitsInside(wrapped, s, radius))
				) {
					outsideW = Math.max(outsideW, size.width);
					outsideH = Math.max(outsideH, size.height);
				}
			}
		}
	}
	if (outsideW > 0) {
		radius = radiusFor(Math.min(box.width / 2 - outsideW - GAP, box.height / 2 - outsideH - GAP));
	}
	const hole = doughnut ? (Math.max(10, Math.min(90, plot.holeSize ?? 50)) / 100) * radius : 0;
	const ringWidth = (radius - hole) / prepared.length;
	const out: ChartRenderNode[] = [];
	prepared.forEach((r, ringIndex) => {
		const meta = { plot: plotIndex, series: r.index };
		const nodes: ChartRenderNode[] = [];
		const innerR = doughnut ? hole + ringIndex * ringWidth : 0;
		const outerR = doughnut ? innerR + ringWidth : radius;
		for (const s of r.slices) {
			const mid = (s.start + s.end) / 2;
			const [ox, oy] = polar(cx, cy, (radius * s.explosion) / 100, mid);
			const paint = shapePaint(ctx, plot, r.series, s.point, r.data.values.length, {
				color: '#FFFFFF',
				width: 1,
			});
			nodes.push({
				type: 'path',
				d: sectorPath(ox, oy, outerR, innerR, s.start, s.end),
				...paint,
				role: 'slice',
				...meta,
				point: s.point,
			});
			const lbl = effectiveLabels(plot, r.series, s.point);
			if (!lbl) {
				continue;
			}
			const text = labelText(ctx, lbl, {
				value: s.value,
				category: r.data.categories[s.point] ?? String(s.point + 1),
				seriesName: r.data.name,
				percent: s.share,
				valueFormat: r.data.valueFormat,
			});
			if (text) {
				labels.push(
					...sliceLabel(ctx, plot, lbl, text, ox, oy, innerR, outerR, s, doughnut, {
						...meta,
						point: s.point,
					}),
				);
			}
		}
		out.push({ type: 'group', children: nodes, role: 'series', ...meta });
	});
	return out;
}

function sliceLabel(
	ctx: LayoutContext,
	plot: ChartPlot,
	lbl: ChartDataLabels,
	text: string,
	cx: number,
	cy: number,
	innerR: number,
	outerR: number,
	slice: PieSlice,
	doughnut: boolean,
	meta: { plot: number; series: number; point: number },
): ChartRenderNode[] {
	const style = labelTextStyle(ctx, plot, lbl);
	const size = labelSize(ctx, style, text);
	const mid = (slice.start + slice.end) / 2;
	const position = doughnut ? 'ctr' : (lbl.position ?? 'bestFit');
	const inside = (fraction: number) => {
		const [x, y] = polar(cx, cy, innerR + (outerR - innerR) * fraction, mid);
		return placeLabel(ctx, lbl, style, text, x, y, 'middle', 'middle', meta);
	};
	const outside = () => {
		const [x, y] = polar(cx, cy, outerR + GAP, mid);
		const a = ((mid % 360) + 360) % 360;
		const vAlign = a > 100 && a < 260 ? 'top' : a < 80 || a > 280 ? 'bottom' : 'middle';
		return placeLabel(ctx, lbl, style, text, x, y, labelAnchor(mid), vAlign, meta);
	};
	switch (position) {
		case 'ctr':
			return inside(doughnut ? 0.5 : 0.5);
		case 'inEnd':
			return inside(0.78);
		case 'outEnd':
			return outside();
		default: {
			// Best fit: inside when the label fits the slice, otherwise just outside it.
			if (fitsInside(size, slice, outerR)) {
				return inside(0.62);
			}
			// Excel breaks a multi-part label at its separator before moving it outside.
			const wrapped = wrapAtSeparator(text, lbl);
			if (wrapped !== text && fitsInside(labelSize(ctx, style, wrapped), slice, outerR)) {
				const [x, y] = polar(cx, cy, innerR + (outerR - innerR) * 0.62, mid);
				return placeLabel(ctx, lbl, style, wrapped, x, y, 'middle', 'middle', meta);
			}
			return outside();
		}
	}
}

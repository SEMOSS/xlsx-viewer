/** Path builders: polylines, smoothed curves, pie / doughnut sectors and markers. */
import type { ChartMarker } from '../../types/chart';
import type { ChartPaint, ChartRenderNode } from './types';

export type Point = [number, number];

const r2 = (v: number): number => Math.round(v * 100) / 100;

/** `M x y L x y ...` through the points. */
export function linePath(points: readonly Point[]): string {
	return points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${r2(x)} ${r2(y)}`).join(' ');
}

/**
 * Smooth curve through the points (Catmull-Rom converted to cubic Bezier
 * segments), close to Excel's "smoothed line".
 */
export function smoothPath(points: readonly Point[]): string {
	if (points.length < 3) {
		return linePath(points);
	}
	const t = 1 / 6;
	let d = `M${r2(points[0][0])} ${r2(points[0][1])}`;
	for (let i = 0; i < points.length - 1; i++) {
		const p0 = points[i - 1] ?? points[i];
		const p1 = points[i];
		const p2 = points[i + 1];
		const p3 = points[i + 2] ?? p2;
		const c1: Point = [p1[0] + (p2[0] - p0[0]) * t, p1[1] + (p2[1] - p0[1]) * t];
		const c2: Point = [p2[0] - (p3[0] - p1[0]) * t, p2[1] - (p3[1] - p1[1]) * t];
		d += ` C${r2(c1[0])} ${r2(c1[1])} ${r2(c2[0])} ${r2(c2[1])} ${r2(p2[0])} ${r2(p2[1])}`;
	}
	return d;
}

/** Point on a circle; `angle` in degrees clockwise from 12 o'clock. */
export function polar(cx: number, cy: number, r: number, angle: number): Point {
	const rad = (angle * Math.PI) / 180;
	return [cx + r * Math.sin(rad), cy - r * Math.cos(rad)];
}

/**
 * Pie slice (inner radius 0) or doughnut segment between two angles
 * (degrees clockwise from 12 o'clock).
 */
export function sectorPath(
	cx: number,
	cy: number,
	outer: number,
	inner: number,
	start: number,
	end: number,
): string {
	const sweep = end - start;
	if (sweep >= 359.999) {
		// Full ring: two half arcs (a single arc cannot describe a full circle).
		const ring = (r: number, dir: 0 | 1) => {
			const [x0, y0] = polar(cx, cy, r, 0);
			const [x1, y1] = polar(cx, cy, r, 180);
			return `M${r2(x0)} ${r2(y0)} A${r2(r)} ${r2(r)} 0 1 ${dir} ${r2(x1)} ${r2(y1)} A${r2(r)} ${r2(r)} 0 1 ${dir} ${r2(x0)} ${r2(y0)} Z`;
		};
		return inner > 0 ? `${ring(outer, 1)} ${ring(inner, 0)}` : ring(outer, 1);
	}
	const large = sweep > 180 ? 1 : 0;
	const [ox0, oy0] = polar(cx, cy, outer, start);
	const [ox1, oy1] = polar(cx, cy, outer, end);
	if (inner <= 0) {
		return `M${r2(cx)} ${r2(cy)} L${r2(ox0)} ${r2(oy0)} A${r2(outer)} ${r2(outer)} 0 ${large} 1 ${r2(ox1)} ${r2(oy1)} Z`;
	}
	const [ix1, iy1] = polar(cx, cy, inner, end);
	const [ix0, iy0] = polar(cx, cy, inner, start);
	return (
		`M${r2(ox0)} ${r2(oy0)} A${r2(outer)} ${r2(outer)} 0 ${large} 1 ${r2(ox1)} ${r2(oy1)} ` +
		`L${r2(ix1)} ${r2(iy1)} A${r2(inner)} ${r2(inner)} 0 ${large} 0 ${r2(ix0)} ${r2(iy0)} Z`
	);
}

/** Marker shape centred on (cx, cy); `size` is the marker box in pixels. */
export function markerNode(
	symbol: NonNullable<ChartMarker['symbol']>,
	cx: number,
	cy: number,
	size: number,
	paint: ChartPaint,
): ChartRenderNode | undefined {
	const h = size / 2;
	const strokeOnly: ChartPaint = {
		stroke: paint.stroke ?? paint.fill,
		strokeWidth: Math.max(1, paint.strokeWidth ?? 1),
	};
	switch (symbol) {
		case 'none':
			return undefined;
		case 'circle':
			return { type: 'circle', cx, cy, r: h, ...paint, role: 'marker' };
		case 'dot':
			return { type: 'circle', cx, cy, r: Math.max(1, h / 2), ...paint, role: 'marker' };
		case 'square':
		case 'picture':
			return {
				type: 'rect',
				x: cx - h,
				y: cy - h,
				width: size,
				height: size,
				...paint,
				role: 'marker',
			};
		case 'dash':
			return {
				type: 'rect',
				x: cx - h,
				y: cy - size / 8,
				width: size,
				height: size / 4,
				...paint,
				role: 'marker',
			};
		case 'diamond':
			return {
				type: 'path',
				d: `M${r2(cx)} ${r2(cy - h)} L${r2(cx + h)} ${r2(cy)} L${r2(cx)} ${r2(cy + h)} L${r2(cx - h)} ${r2(cy)} Z`,
				...paint,
				role: 'marker',
			};
		case 'triangle':
			return {
				type: 'path',
				d: `M${r2(cx)} ${r2(cy - h)} L${r2(cx + h)} ${r2(cy + h)} L${r2(cx - h)} ${r2(cy + h)} Z`,
				...paint,
				role: 'marker',
			};
		case 'x':
			return {
				type: 'path',
				d: `M${r2(cx - h)} ${r2(cy - h)} L${r2(cx + h)} ${r2(cy + h)} M${r2(cx + h)} ${r2(cy - h)} L${r2(cx - h)} ${r2(cy + h)}`,
				...strokeOnly,
				role: 'marker',
			};
		case 'plus':
			return {
				type: 'path',
				d: `M${r2(cx)} ${r2(cy - h)} L${r2(cx)} ${r2(cy + h)} M${r2(cx - h)} ${r2(cy)} L${r2(cx + h)} ${r2(cy)}`,
				...strokeOnly,
				role: 'marker',
			};
		case 'star':
			return {
				type: 'path',
				d:
					`M${r2(cx - h)} ${r2(cy - h)} L${r2(cx + h)} ${r2(cy + h)} M${r2(cx + h)} ${r2(cy - h)} ` +
					`L${r2(cx - h)} ${r2(cy + h)} M${r2(cx)} ${r2(cy - h)} L${r2(cx)} ${r2(cy + h)}`,
				...strokeOnly,
				role: 'marker',
			};
		default:
			return { type: 'circle', cx, cy, r: h, ...paint, role: 'marker' };
	}
}

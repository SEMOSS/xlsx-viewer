/**
 * Render tree produced by the chart layout engine: framework-neutral drawing
 * primitives in pixel coordinates (origin top-left), ready to be mapped onto
 * SVG, canvas or any other 2-D surface.
 */

export interface ChartGradient {
	kind: 'linear' | 'radial';
	/** Linear gradients: direction in degrees (0 = left to right, 90 = top to bottom). */
	angle: number;
	/** Offsets in [0, 1]; colors are CSS colors. */
	stops: Array<{ offset: number; color: string }>;
}

/** Fill and stroke shared by the shape primitives. Absent `fill` / `stroke` means none. */
export interface ChartPaint {
	/** CSS color. */
	fill?: string;
	/** Gradient fill; takes precedence over `fill`. */
	gradient?: ChartGradient;
	stroke?: string;
	strokeWidth?: number;
	/** Dash pattern in pixels. */
	dash?: number[];
	lineCap?: 'butt' | 'round' | 'square';
	lineJoin?: 'miter' | 'round' | 'bevel';
	opacity?: number;
}

export type ChartNodeRole =
	| 'chart-area'
	| 'plot-area'
	| 'title'
	| 'legend'
	| 'legend-entry'
	| 'legend-key'
	| 'axis'
	| 'axis-line'
	| 'tick'
	| 'tick-label'
	| 'axis-title'
	| 'gridline'
	| 'series'
	| 'bar'
	| 'line'
	| 'area'
	| 'marker'
	| 'slice'
	| 'bubble'
	| 'hi-low-line'
	| 'up-down-bar'
	| 'trendline'
	| 'data-label'
	| 'placeholder';

/** Identification attached to nodes so hosts can add interactivity or tests can find them. */
export interface ChartNodeMeta {
	role?: ChartNodeRole;
	/** Index into `chart.plots`. */
	plot?: number;
	/** Index into `plot.series`. */
	series?: number;
	/** Point (category) index. */
	point?: number;
}

export interface ChartRectNode extends ChartPaint, ChartNodeMeta {
	type: 'rect';
	x: number;
	y: number;
	width: number;
	height: number;
	/** Corner radius. */
	rx?: number;
}

export interface ChartPathNode extends ChartPaint, ChartNodeMeta {
	type: 'path';
	/** SVG path data. */
	d: string;
}

export interface ChartLineNode extends ChartPaint, ChartNodeMeta {
	type: 'line';
	x1: number;
	y1: number;
	x2: number;
	y2: number;
}

export interface ChartPolylineNode extends ChartPaint, ChartNodeMeta {
	type: 'polyline';
	/** Flat list of coordinates: `[x0, y0, x1, y1, ...]`. */
	points: number[];
}

export interface ChartCircleNode extends ChartPaint, ChartNodeMeta {
	type: 'circle';
	cx: number;
	cy: number;
	r: number;
}

/** One formatted run of a rich text line. */
export interface ChartTextRun {
	text: string;
	fontSize?: number;
	fontWeight?: 'normal' | 'bold';
	fontStyle?: 'normal' | 'italic';
	decoration?: 'underline' | 'line-through';
	fill?: string;
	fontFamily?: string;
}

export interface ChartTextNode extends ChartNodeMeta {
	type: 'text';
	/** Anchor x (see `anchor`). */
	x: number;
	/** Baseline of the first line. */
	y: number;
	lines: string[];
	/** Rich runs per line (titles with mixed formatting); overrides `lines` when present. */
	runs?: ChartTextRun[][];
	/** Distance between baselines. */
	lineHeight: number;
	fontSize: number;
	fontFamily: string;
	fontWeight?: 'normal' | 'bold';
	fontStyle?: 'normal' | 'italic';
	decoration?: 'underline' | 'line-through';
	fill: string;
	anchor: 'start' | 'middle' | 'end';
	/** Clockwise rotation in degrees around (`pivotX`, `pivotY`). */
	rotation?: number;
	pivotX?: number;
	pivotY?: number;
}

export interface ChartGroupNode extends ChartNodeMeta {
	type: 'group';
	children: ChartRenderNode[];
	/** Clip children to this rectangle. */
	clip?: { x: number; y: number; width: number; height: number };
}

export type ChartRenderNode =
	| ChartRectNode
	| ChartPathNode
	| ChartLineNode
	| ChartPolylineNode
	| ChartCircleNode
	| ChartTextNode
	| ChartGroupNode;

export interface ChartRenderTree {
	width: number;
	height: number;
	root: ChartGroupNode;
	/** Plain-language summary (title and chart type), for accessibility. */
	description: string;
}

/** An axis-aligned rectangle. */
export interface ChartBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

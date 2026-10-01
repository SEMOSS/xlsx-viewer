import type {
	DrawingColor,
	DrawingParagraph,
	DrawingTextStyle,
	ShapeFill,
	ShapeLine,
} from './drawing';

/**
 * Chart model parsed from a DrawingML chart part (`xl/charts/chartN.xml`).
 * Series reference worksheet ranges (`ref`) and also carry the cached values
 * stored in the file, so a chart can render before (or without) recalculation.
 */

export type ChartPlotType =
	| 'bar'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'radar'
	| 'bubble'
	| 'stock'
	| 'surface'
	| 'ofPie'
	| 'unknown';

/**
 * Manual position of a chart element (`c:manualLayout`). Values are fractions
 * of the chart area; `edge` positions are absolute, `factor` ones are offsets
 * from the automatic position.
 */
export interface ChartManualLayout {
	/** Plot area only: the box covers the inner plot (`inner`) or includes tick labels (`outer`). */
	layoutTarget?: 'inner' | 'outer';
	xMode?: 'edge' | 'factor';
	yMode?: 'edge' | 'factor';
	x?: number;
	y?: number;
	w?: number;
	h?: number;
}

export interface ChartText {
	/** Static text (or the cached text of `ref`). */
	text?: string;
	/** Formula reference for the text, e.g. `Sheet1!$B$1`. */
	ref?: string;
	style?: DrawingTextStyle;
	overlay?: boolean;
	/** Rich text (`c:rich`) paragraphs with per-run formatting; `text` holds the plain text. */
	paragraphs?: DrawingParagraph[];
	/** Text rotation in degrees (`a:bodyPr/@rot`), e.g. -90 for a vertical axis title. */
	rotation?: number;
	layout?: ChartManualLayout;
}

export interface ChartDataRef {
	/** Formula reference, e.g. `Sheet1!$A$2:$A$6`. */
	ref?: string;
	/** Cached values from the file, by point index (gaps are `null`). */
	cache?: Array<string | number | null>;
	formatCode?: string;
	/** Cached multi-level category labels (outer levels first). */
	levels?: Array<Array<string | null>>;
	kind: 'num' | 'str' | 'multiLvlStr' | 'literal';
}

export interface ChartMarker {
	symbol?:
		| 'none'
		| 'auto'
		| 'circle'
		| 'dash'
		| 'diamond'
		| 'dot'
		| 'picture'
		| 'plus'
		| 'square'
		| 'star'
		| 'triangle'
		| 'x';
	size?: number;
	fill?: ShapeFill;
	line?: ShapeLine;
}

export interface ChartDataLabels {
	showValue?: boolean;
	showCategoryName?: boolean;
	showSeriesName?: boolean;
	showPercent?: boolean;
	showLegendKey?: boolean;
	position?: 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';
	numFmt?: string;
	separator?: string;
	style?: DrawingTextStyle;
	deleted?: boolean;
	showBubbleSize?: boolean;
	showLeaderLines?: boolean;
	/** Label box fill and border (`c:spPr`). */
	fill?: ShapeFill;
	line?: ShapeLine;
	/** Per-point overrides (`c:dLbl`), e.g. a single deleted or repositioned label. */
	points?: Array<ChartDataLabels & { idx: number; layout?: ChartManualLayout }>;
}

export interface ChartDataPoint {
	idx: number;
	fill?: ShapeFill;
	line?: ShapeLine;
	/** Pie slice explosion (percent of radius). */
	explosion?: number;
	marker?: ChartMarker;
	invertIfNegative?: boolean;
}

export interface ChartTrendline {
	type: 'linear' | 'exp' | 'log' | 'movingAvg' | 'poly' | 'power';
	order?: number;
	period?: number;
	line?: ShapeLine;
	name?: string;
}

export interface ChartSeries {
	idx: number;
	order: number;
	name?: ChartText;
	/** Categories (`c:cat`) or X values (`c:xVal`) for scatter and bubble charts. */
	categories?: ChartDataRef;
	/** Values (`c:val`) or Y values (`c:yVal`). */
	values?: ChartDataRef;
	bubbleSizes?: ChartDataRef;
	fill?: ShapeFill;
	line?: ShapeLine;
	marker?: ChartMarker;
	points?: ChartDataPoint[];
	dataLabels?: ChartDataLabels;
	smooth?: boolean;
	explosion?: number;
	invertIfNegative?: boolean;
	trendlines?: ChartTrendline[];
}

export interface ChartPlot {
	type: ChartPlotType;
	/** Bar charts: horizontal bars (`bar`) or vertical columns (`col`). */
	barDirection?: 'bar' | 'col';
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	varyColors?: boolean;
	series: ChartSeries[];
	/** Ids of the axes this plot uses (category/X axis first). */
	axisIds: number[];
	gapWidth?: number;
	overlap?: number;
	/** Doughnut hole size in percent. */
	holeSize?: number;
	/** Pie first slice angle in degrees. */
	firstSliceAngle?: number;
	scatterStyle?: 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker' | 'none';
	radarStyle?: 'standard' | 'marker' | 'filled';
	dataLabels?: ChartDataLabels;
	/** Line charts: draw markers. */
	marker?: boolean;
	is3D?: boolean;
	/** Stock / line charts: high-low lines, up/down bars, drop lines. */
	hiLowLines?: boolean;
	upDownBars?: boolean;
	/** Bubble charts: size of the largest bubble in percent of the default (`c:bubbleScale`). */
	bubbleScale?: number;
	/** Bubble charts: the size value maps to the bubble area (default) or its width. */
	bubbleSizeRepresents?: 'area' | 'w';
}

export interface ChartAxis {
	id: number;
	kind: 'cat' | 'val' | 'date' | 'ser';
	position: 'b' | 'l' | 'r' | 't';
	crossAxisId: number;
	deleted?: boolean;
	title?: ChartText;
	numFmt?: { code: string; sourceLinked?: boolean };
	min?: number;
	max?: number;
	majorUnit?: number;
	minorUnit?: number;
	logBase?: number;
	orientation?: 'minMax' | 'maxMin';
	crosses?: 'autoZero' | 'min' | 'max';
	crossesAt?: number;
	majorGridlines?: ShapeLine | true;
	minorGridlines?: ShapeLine | true;
	majorTickMark?: 'cross' | 'in' | 'none' | 'out';
	minorTickMark?: 'cross' | 'in' | 'none' | 'out';
	tickLabelPosition?: 'high' | 'low' | 'nextTo' | 'none';
	line?: ShapeLine;
	textStyle?: DrawingTextStyle;
	/** Category axis: put values between tick marks (`crossBetween="between"`). */
	crossBetween?: 'between' | 'midCat';
	/** Tick label rotation in degrees (`c:txPr/a:bodyPr/@rot`); omitted means automatic. */
	labelRotation?: number;
	/** Category / series axis: show every n-th tick label (`c:tickLblSkip`). */
	tickLabelSkip?: number;
	/** Value axis display units (`c:dispUnits`): a built-in unit name or a custom divisor. */
	displayUnits?: { unit: string | number; showLabel?: boolean };
	/** Date axis time units (`c:baseTimeUnit`, `c:majorTimeUnit`). */
	baseTimeUnit?: 'days' | 'months' | 'years';
	majorTimeUnit?: 'days' | 'months' | 'years';
	/** Category axis: draw only the innermost level of multi-level labels (`c:noMultiLvlLbl`). */
	noMultiLevelLabels?: boolean;
}

export interface ChartLegend {
	position: 'r' | 'l' | 't' | 'b' | 'tr';
	overlay?: boolean;
	/** Indices of legend entries that were deleted. */
	deletedEntries?: number[];
	textStyle?: DrawingTextStyle;
	fill?: ShapeFill;
	line?: ShapeLine;
	layout?: ChartManualLayout;
}

export interface ChartModel {
	title?: ChartText;
	autoTitleDeleted?: boolean;
	plots: ChartPlot[];
	axes: ChartAxis[];
	legend?: ChartLegend;
	plotVisibleOnly?: boolean;
	displayBlanksAs?: 'gap' | 'zero' | 'span';
	/** Built-in chart style number (`c:style`). */
	style?: number;
	roundedCorners?: boolean;
	/** Chart area (outer) fill and border. */
	chartFill?: ShapeFill;
	chartLine?: ShapeLine;
	/** Plot area fill and border. */
	plotFill?: ShapeFill;
	plotLine?: ShapeLine;
	/** Manual plot area position (`c:plotArea/c:layout`). */
	plotAreaLayout?: ChartManualLayout;
	/** Default text style for the whole chart (`c:txPr`). */
	textStyle?: DrawingTextStyle;
	/** Original chart part XML, re-emitted when the chart is not edited. */
	rawXml?: string;
	/** Set when the chart was edited or created and its part must be regenerated. */
	dirty?: boolean;
}

/** Convenience: an explicit color used by renderers after theme resolution. */
export type ResolvedChartColor = string;

export type { DrawingColor };

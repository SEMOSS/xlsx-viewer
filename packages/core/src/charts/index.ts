/**
 * Charts: DrawingML chart parts <-> ChartModel, live data resolution, chart
 * creation from a cell range and a framework-neutral layout engine that turns
 * a chart into drawing primitives.
 */
export { parseChartXml } from './parse/parse-chart';
export type { ChartParseContext } from './parse/parse-chart';
export { writeChartXml } from './write/write-chart';

export { resolveChartData, cachedChartData } from './data/resolve-data';
export { resolveChartRef, formatChartRef } from './data/chart-ref';
export type { ChartRefArea } from './data/chart-ref';
export type { ChartData, ChartPlotData, ChartSeriesData } from './data/types';
export { createChartFromRange } from './data/create-chart';
export type { CreatableChartType, CreateChartOptions } from './data/create-chart';

export { layoutChart } from './layout/layout-chart';
export type { ChartLayoutOptions } from './layout/layout-chart';
export { describeChart, chartTitleText } from './layout/describe';
export { computeValueScale, niceStep } from './layout/scale';
export type { ValueScale, ValueScaleOptions } from './layout/scale';
export { autoChartColor, autoMarkerSymbol, chartStyleIndex } from './layout/palette';
export { defaultMeasureText as defaultChartTextMeasurer } from './layout/text';
export type { ChartFont, ChartTextMeasurer } from './layout/text';
export type {
	ChartBox,
	ChartCircleNode,
	ChartGradient,
	ChartGroupNode,
	ChartLineNode,
	ChartNodeMeta,
	ChartNodeRole,
	ChartPaint,
	ChartPathNode,
	ChartPolylineNode,
	ChartRectNode,
	ChartRenderNode,
	ChartRenderTree,
	ChartTextNode,
	ChartTextRun,
} from './layout/types';

export {
	applyChartStyle,
	applyColorPalette,
	applyQuickLayout,
	CHART_COLOR_PALETTES,
	CHART_QUICK_LAYOUT_COUNT,
	CHART_QUICK_STYLE_COUNT,
	CHART_TYPE_TARGETS,
	chartElementState,
	chartPaletteColor,
	chartTypeTargetOf,
	convertChartType,
	setChartElement,
	switchRowColumn,
} from './edit';
export type {
	ChartColorPalette,
	ChartElementKind,
	ChartElementOptions,
	ChartTypeTarget,
} from './edit';

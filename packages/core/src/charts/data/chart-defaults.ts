/**
 * Explicit formatting Excel 2016+ writes for newly inserted charts (chart
 * style 2 / "Style 1"), so charts created here look the same in Excel as in
 * the viewer.
 */
import type { ChartAxis, ChartSeries } from '../../types/chart';
import type { DrawingColor, DrawingTextStyle, ShapeFill, ShapeLine } from '../../types/drawing';
import { autoChartColor } from '../layout/palette';

const tx1 = (lumMod: number, lumOff: number): DrawingColor => ({
	kind: 'scheme',
	value: 'tx1',
	transforms: [
		{ name: 'lumMod', val: lumMod },
		{ name: 'lumOff', val: lumOff },
	],
});

/** Text 1 at 65%: #595959 on the Office theme. */
export const TEXT_COLOR = tx1(65000, 35000);
/** Text 1 at 15%: #D9D9D9 (gridlines, chart border, category axis). */
export const GRID_COLOR = tx1(15000, 85000);
/** Text 1 at 25%: #BFBFBF (value axis lines of XY charts). */
export const AXIS_COLOR = tx1(25000, 75000);

export const NO_FILL: ShapeFill = { type: 'none' };
export const NO_LINE: ShapeLine = { fill: NO_FILL };

export function solid(color: DrawingColor): ShapeFill {
	return { type: 'solid', color };
}

export function hairline(color: DrawingColor): ShapeLine {
	return { width: 9525, cap: 'flat', fill: solid(color) };
}

export function labelText(size: number): DrawingTextStyle {
	return { size, bold: false, color: TEXT_COLOR };
}

/** Category axis (or the X value axis of XY charts). */
export function defaultCategoryAxis(
	id: number,
	crossAxisId: number,
	position: ChartAxis['position'],
	kind: ChartAxis['kind'] = 'cat',
): ChartAxis {
	const axis: ChartAxis = {
		id,
		kind,
		position,
		crossAxisId,
		deleted: false,
		numFmt: { code: 'General', sourceLinked: true },
		orientation: 'minMax',
		majorTickMark: 'none',
		minorTickMark: 'none',
		tickLabelPosition: 'nextTo',
		line: hairline(kind === 'val' ? AXIS_COLOR : GRID_COLOR),
		textStyle: labelText(9),
		crosses: 'autoZero',
	};
	if (kind === 'val') {
		axis.majorGridlines = hairline(GRID_COLOR);
		axis.crossBetween = 'midCat';
	}
	return axis;
}

/** Value axis with major gridlines and no axis line. */
export function defaultValueAxis(
	id: number,
	crossAxisId: number,
	position: ChartAxis['position'],
	crossBetween: 'between' | 'midCat',
): ChartAxis {
	return {
		id,
		kind: 'val',
		position,
		crossAxisId,
		deleted: false,
		majorGridlines: hairline(GRID_COLOR),
		numFmt: { code: 'General', sourceLinked: true },
		orientation: 'minMax',
		majorTickMark: 'none',
		minorTickMark: 'none',
		tickLabelPosition: 'nextTo',
		line: NO_LINE,
		textStyle: labelText(9),
		crosses: 'autoZero',
		crossBetween,
	};
}

export type CreatableChartType =
	| 'column'
	| 'bar'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'radar';

/** Series formatting for a chart type (`index` picks the automatic accent). */
export function styleSeries(
	series: ChartSeries,
	type: CreatableChartType,
	index: number,
	count: number,
): void {
	const color = autoChartColor(index, count, 2);
	switch (type) {
		case 'column':
		case 'bar':
		case 'area':
			series.fill = solid(color);
			series.line = NO_LINE;
			break;
		case 'line':
		case 'radar':
			series.line = { width: 28575, cap: 'rnd', fill: solid(color) };
			series.marker = { symbol: 'none' };
			break;
		case 'scatter':
			series.line = { width: 25400, cap: 'rnd', fill: NO_FILL };
			series.marker = {
				symbol: 'circle',
				size: 5,
				fill: solid(color),
				line: { width: 9525, fill: solid(color) },
			};
			break;
		case 'pie':
		case 'doughnut':
			series.line = { width: 19050, fill: solid({ kind: 'scheme', value: 'bg1' }) };
			break;
	}
}

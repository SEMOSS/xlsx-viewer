/**
 * `<ChartView>`: draws a worksheet chart as crisp, scalable inline SVG. Data is
 * resolved live from the workbook's cells and laid out by the core chart
 * engine; both steps are memoized.
 */
import { useMemo } from 'react';
import type { CSSProperties } from 'react';
import {
	type ChartModel,
	type ChartTextMeasurer,
	type Workbook,
	layoutChart,
	resolveChartData,
} from 'xlsx-viewer-core';
import { ChartSvg } from './ChartSvg';
import { canvasTextMeasurer } from './text-measure';

export interface ChartViewProps {
	chart: ChartModel;
	workbook: Workbook;
	/** Sheet holding the chart (scopes sheet-level names and unqualified references). */
	sheetIndex: number;
	/** Size in CSS pixels. */
	width: number;
	height: number;
	className?: string;
	style?: CSSProperties;
	/**
	 * Change this (e.g. a workbook revision counter) after editing cells: the
	 * workbook object itself stays the same when cells change, so the chart
	 * re-reads its source data only when this value changes.
	 */
	revision?: number | string;
	/** Text measurer override; defaults to canvas measurement when available. */
	measureText?: ChartTextMeasurer;
}

export function ChartView({
	chart,
	workbook,
	sheetIndex,
	width,
	height,
	className,
	style,
	revision,
	measureText,
}: ChartViewProps) {
	const data = useMemo(
		() => resolveChartData(chart, workbook, sheetIndex),
		// `revision` invalidates the cached data after in-place cell edits.
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[chart, workbook, sheetIndex, revision],
	);
	const measure = measureText ?? canvasTextMeasurer();
	const tree = useMemo(
		() =>
			layoutChart(chart, data, workbook.theme, Math.max(1, width), Math.max(1, height), {
				measureText: measure,
			}),
		[chart, data, workbook.theme, width, height, measure],
	);
	return <ChartSvg tree={tree} className={className} style={style} />;
}

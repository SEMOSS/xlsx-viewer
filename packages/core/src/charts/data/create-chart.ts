/**
 * Insert-chart helper: build a ChartModel from a cell range with Excel's
 * heuristics (header row -> series names, leading text column(s) ->
 * categories, series in columns unless the data is wider than tall).
 */
import { cellKey } from '../../address/cell-address';
import { formatValue, isDateFormat } from '../../numfmt';
import { StyleResolver } from '../../styles/resolve-style';
import type { RangeAddress } from '../../types/address';
import type { CellValue } from '../../types/cell';
import type {
	ChartAxis,
	ChartDataRef,
	ChartLegend,
	ChartModel,
	ChartPlot,
	ChartSeries,
} from '../../types/chart';
import type { Workbook } from '../../types/workbook';
import {
	type CreatableChartType,
	GRID_COLOR,
	defaultCategoryAxis,
	defaultValueAxis,
	hairline,
	labelText,
	solid,
	styleSeries,
} from './chart-defaults';
import { formatChartRef } from './chart-ref';
import { currentRegion, usedBounds } from './source-region';

export type { CreatableChartType } from './chart-defaults';

export interface CreateChartOptions {
	/** Title text; omitted: Excel's automatic title (the series name, or "Chart Title"); `false`: none. */
	title?: string | false;
	/** Plot series from columns or rows. Default: columns unless the data has more columns than rows. */
	seriesIn?: 'columns' | 'rows';
	/** Grouping for column, bar, line and area charts. */
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	/** Legend position, or `false` for none. Default: bottom, except single-series cartesian charts. */
	legend?: ChartLegend['position'] | false;
	/** First axis id; consecutive ids are used. */
	axisIdBase?: number;
}

interface Cell {
	value: CellValue;
	text: string;
	format: string;
}

const isEmpty = (c: Cell): boolean => c.value === null || c.value === '';
const isNumber = (c: Cell): boolean => typeof c.value === 'number';
const isLabel = (c: Cell): boolean => !isEmpty(c) && !isNumber(c);

/** Number of leading label columns (text or dates) of the data rows; outer levels first. */
function labelColumns(grid: Cell[][], firstRow: number, dates: (c: Cell) => boolean): number {
	const width = grid[0]?.length ?? 0;
	let n = 0;
	for (let c = 0; c < width - 1; c++) {
		const cells = grid.slice(firstRow).map((row) => row[c]);
		const labels = cells.every(
			(cell) => isEmpty(cell) || isLabel(cell) || (n === 0 && dates(cell)),
		);
		if (!labels || !cells.some((cell) => !isEmpty(cell))) {
			break;
		}
		n++;
		if (cells.some(dates)) {
			break;
		}
	}
	return n;
}

function transpose(grid: Cell[][]): Cell[][] {
	return (grid[0] ?? []).map((_, c) => grid.map((row) => row[c]));
}

/**
 * Create a chart for `range` on sheet `sheetIndex`. The model is marked
 * `dirty` so the writer serializes it; its caches hold the current values.
 */
export function createChartFromRange(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	type: CreatableChartType,
	options: CreateChartOptions = {},
): ChartModel {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) {
		throw new RangeError(`No sheet at index ${sheetIndex}`);
	}
	let area: RangeAddress = { ...range };
	if (area.top === area.bottom && area.left === area.right) {
		area = currentRegion(sheet, area);
	}
	const used = usedBounds(sheet);
	if (used) {
		area = {
			top: Math.max(area.top, Math.min(area.bottom, used.top)),
			left: Math.max(area.left, Math.min(area.right, used.left)),
			bottom: Math.min(area.bottom, Math.max(area.top, used.bottom)),
			right: Math.min(area.right, Math.max(area.left, used.right)),
		};
	}
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const read = (r: number, c: number): Cell => {
		const cell = sheet.cells.get(cellKey(r, c));
		const value = cell?.value ?? null;
		const format = styles.resolve(cell?.style).numFmt;
		return {
			value,
			format,
			text: value === null ? '' : formatValue(value, format, { date1904: workbook.date1904 }).text,
		};
	};
	const rows = area.bottom - area.top + 1;
	const cols = area.right - area.left + 1;
	const grid: Cell[][] = Array.from({ length: rows }, (_, r) =>
		Array.from({ length: cols }, (_, c) => read(area.top + r, area.left + c)),
	);
	const isDate = (c: Cell) => isNumber(c) && isDateFormat(c.format);
	// Header row: an empty top-left corner, or labels (no plain numbers) across the top.
	const corner = grid[0][0];
	const firstRow = grid[0];
	const headerRow =
		rows > 1 &&
		(isEmpty(corner) ||
			(firstRow.slice(1).some(isLabel) &&
				firstRow.slice(1).every((c) => !isNumber(c) || isDate(c))));
	let labelCols = labelColumns(grid, headerRow ? 1 : 0, isDate);
	if (labelCols === 0 && headerRow && isEmpty(corner) && cols > 1) {
		labelCols = 1;
	}
	const dataRows = rows - (headerRow ? 1 : 0);
	const dataCols = cols - labelCols;
	const inColumns = options.seriesIn ? options.seriesIn === 'columns' : dataRows >= dataCols;
	return buildModel(sheet.name, area, grid, {
		type,
		options,
		inColumns,
		headerRow,
		labelCols,
		xy: type === 'scatter',
	});
}

interface Plan {
	type: CreatableChartType;
	options: CreateChartOptions;
	inColumns: boolean;
	headerRow: boolean;
	labelCols: number;
	xy: boolean;
}

function numRef(sheetName: string, range: RangeAddress, cells: Cell[]): ChartDataRef {
	const firstNumber = cells.find(isNumber);
	return {
		kind: 'num',
		ref: formatChartRef(sheetName, range),
		cache: cells.map((c) => (typeof c.value === 'number' ? c.value : null)),
		formatCode: firstNumber && firstNumber.format !== 'General' ? firstNumber.format : 'General',
	};
}

function buildModel(
	sheetName: string,
	area: RangeAddress,
	cellGrid: Cell[][],
	plan: Plan,
): ChartModel {
	// Work in "series in columns" orientation: rows are points, columns are series.
	const grid = plan.inColumns ? cellGrid : transpose(cellGrid);
	const nameRows = plan.inColumns ? (plan.headerRow ? 1 : 0) : plan.labelCols;
	let catCols = plan.inColumns ? plan.labelCols : plan.headerRow ? 1 : 0;
	const at = (r: number, c: number): RangeAddress =>
		plan.inColumns
			? { top: area.top + r, left: area.left + c, bottom: area.top + r, right: area.left + c }
			: { top: area.top + c, left: area.left + r, bottom: area.top + c, right: area.left + r };
	const span = (r0: number, c0: number, r1: number, c1: number): RangeAddress => {
		const a = at(r0, c0);
		const b = at(r1, c1);
		return {
			top: Math.min(a.top, b.top),
			left: Math.min(a.left, b.left),
			bottom: Math.max(a.bottom, b.bottom),
			right: Math.max(a.right, b.right),
		};
	};
	const width = grid[0]?.length ?? 0;
	const lastRow = grid.length - 1;
	// Scatter charts use the first numeric column as X values when there is more than one.
	if (plan.xy && catCols === 0 && width - catCols > 1) {
		catCols = 1;
	}
	let categories: ChartDataRef | undefined;
	if (catCols > 0 && lastRow >= nameRows) {
		const cells = grid.slice(nameRows).map((row) => row.slice(0, catCols));
		const range = span(nameRows, 0, lastRow, catCols - 1);
		const inner = cells.map((row) => row[catCols - 1]);
		if (plan.xy || inner.every((c) => isEmpty(c) || isNumber(c))) {
			categories = catCols === 1 ? numRef(sheetName, range, inner) : undefined;
		}
		if (!categories) {
			categories = {
				kind: catCols > 1 ? 'multiLvlStr' : 'str',
				ref: formatChartRef(sheetName, range),
				cache: inner.map((c) => c.text),
			};
			if (catCols > 1) {
				categories.levels = Array.from({ length: catCols }, (_, l) =>
					cells.map((row) => (isEmpty(row[l]) ? null : row[l].text)),
				);
			}
		}
	}
	const series: ChartSeries[] = [];
	for (let c = catCols; c < width; c++) {
		const idx = series.length;
		const s: ChartSeries = { idx, order: idx };
		if (nameRows > 0) {
			const text = grid
				.slice(0, nameRows)
				.map((row) => row[c].text)
				.filter((t) => t !== '')
				.join(' ');
			s.name = { ref: formatChartRef(sheetName, span(0, c, nameRows - 1, c)), text };
		}
		if (categories) {
			s.categories = { ...categories, cache: [...(categories.cache ?? [])] };
		}
		if (lastRow >= nameRows) {
			s.values = numRef(
				sheetName,
				span(nameRows, c, lastRow, c),
				grid.slice(nameRows).map((row) => row[c]),
			);
		}
		series.push(s);
	}
	series.forEach((s, i) => styleSeries(s, plan.type, i, series.length));
	return assemble(plan, series);
}

function assemble(plan: Plan, series: ChartSeries[]): ChartModel {
	const { type, options } = plan;
	const base = options.axisIdBase ?? 500000001;
	const horizontal = type === 'bar';
	const plot: ChartPlot = { type: 'bar', series, axisIds: [] };
	const axes: ChartAxis[] = [];
	const stackedGrouping = options.grouping === 'stacked' || options.grouping === 'percentStacked';
	switch (type) {
		case 'column':
		case 'bar':
			Object.assign(plot, {
				type: 'bar',
				barDirection: horizontal ? 'bar' : 'col',
				grouping:
					options.grouping && options.grouping !== 'standard' ? options.grouping : 'clustered',
				varyColors: false,
				gapWidth: stackedGrouping ? 150 : 219,
				overlap: stackedGrouping ? 100 : -27,
			});
			break;
		case 'line':
			Object.assign(plot, {
				type: 'line',
				grouping: stackedGrouping ? options.grouping : 'standard',
				varyColors: false,
				marker: true,
			});
			break;
		case 'area':
			Object.assign(plot, {
				type: 'area',
				grouping: stackedGrouping ? options.grouping : 'standard',
				varyColors: false,
			});
			break;
		case 'pie':
			Object.assign(plot, { type: 'pie', varyColors: true, firstSliceAngle: 0 });
			break;
		case 'doughnut':
			Object.assign(plot, { type: 'doughnut', varyColors: true, firstSliceAngle: 0, holeSize: 75 });
			break;
		case 'scatter':
			Object.assign(plot, { type: 'scatter', scatterStyle: 'lineMarker', varyColors: false });
			break;
		case 'radar':
			Object.assign(plot, { type: 'radar', radarStyle: 'marker', varyColors: false });
			break;
	}
	if (type === 'scatter') {
		axes.push(
			defaultCategoryAxis(base, base + 1, 'b', 'val'),
			defaultValueAxis(base + 1, base, 'l', 'midCat'),
		);
	} else if (type === 'radar') {
		const cat = defaultCategoryAxis(base, base + 1, 'b');
		cat.majorGridlines = hairline(GRID_COLOR);
		axes.push(cat, defaultValueAxis(base + 1, base, 'l', 'between'));
	} else if (type !== 'pie' && type !== 'doughnut') {
		axes.push(
			defaultCategoryAxis(base, base + 1, horizontal ? 'l' : 'b'),
			defaultValueAxis(
				base + 1,
				base,
				horizontal ? 'b' : 'l',
				type === 'area' ? 'midCat' : 'between',
			),
		);
	}
	plot.axisIds = axes.map((a) => a.id);
	const pieLike = type === 'pie' || type === 'doughnut';
	const model: ChartModel = {
		plots: [plot],
		axes,
		autoTitleDeleted: options.title === false,
		plotVisibleOnly: true,
		displayBlanksAs: 'gap',
		style: 102,
		roundedCorners: false,
		chartFill: solid({ kind: 'scheme', value: 'bg1' }),
		chartLine: hairline(GRID_COLOR),
		plotFill: { type: 'none' },
		dirty: true,
	};
	if (options.title !== false) {
		model.title = { overlay: false, style: labelText(14) };
		if (options.title !== undefined) {
			model.title.text = options.title;
		}
	}
	const legend = options.legend ?? (pieLike || series.length > 1 ? 'b' : false);
	if (legend) {
		model.legend = { position: legend, overlay: false, textStyle: labelText(9) };
	}
	return model;
}

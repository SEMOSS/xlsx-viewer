/**
 * Resolve chart series against the workbook's current cell values, the way
 * Excel redraws a chart after its source data changes. References that cannot
 * be read (external workbooks, functions, missing sheets) fall back to the
 * values cached in the chart part.
 */
import { cellKey, keyCol, keyRow } from '../../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import { formatValue } from '../../numfmt';
import { StyleResolver } from '../../styles/resolve-style';
import type { CellValue } from '../../types/cell';
import type { ChartDataRef, ChartModel, ChartSeries } from '../../types/chart';
import type { Worksheet } from '../../types/sheet';
import type { Workbook } from '../../types/workbook';
import { isCellError } from '../../values/errors';
import { type ChartRefArea, resolveChartRef } from './chart-ref';
import type { ChartData, ChartSeriesData } from './types';

/** Upper bound on cells read per reference, so a whole-column reference cannot stall rendering. */
const MAX_POINTS = 100_000;

interface ReadCell {
	value: CellValue;
	format: string;
}

/** A visible sub-grid of a reference: rows x cols of cells. */
type Grid = ReadCell[][];

type Blanks = 'gap' | 'zero' | 'span';

class CellSource {
	private readonly styles: StyleResolver;
	private readonly usedBottom = new Map<number, number>();
	private readonly usedRight = new Map<number, number>();

	constructor(
		readonly workbook: Workbook,
		readonly sheetIndex: number,
		readonly visibleOnly: boolean,
	) {
		this.styles = new StyleResolver(workbook.styles, workbook.theme);
	}

	areas(ref: string | undefined): ChartRefArea[] | null {
		if (!ref || /^\s*"/.test(ref)) {
			return null;
		}
		return resolveChartRef(ref, this.workbook, this.sheetIndex);
	}

	private bounds(sheetIndex: number): { bottom: number; right: number } {
		let bottom = this.usedBottom.get(sheetIndex);
		let right = this.usedRight.get(sheetIndex);
		if (bottom === undefined || right === undefined) {
			bottom = 0;
			right = 0;
			for (const key of this.workbook.sheets[sheetIndex].cells.keys()) {
				bottom = Math.max(bottom, keyRow(key));
				right = Math.max(right, keyCol(key));
			}
			this.usedBottom.set(sheetIndex, bottom);
			this.usedRight.set(sheetIndex, right);
		}
		return { bottom, right };
	}

	/** Cells of an area with hidden rows / columns removed (when plotting visible cells only). */
	grid(area: ChartRefArea): Grid {
		const sheet = this.workbook.sheets[area.sheetIndex];
		const { top, left } = area.range;
		let { bottom, right } = area.range;
		if (bottom === MAX_ROW_INDEX || right === MAX_COL_INDEX) {
			const used = this.bounds(area.sheetIndex);
			bottom = Math.min(bottom, used.bottom);
			right = Math.min(right, used.right);
		}
		const rows: number[] = [];
		for (let r = top; r <= bottom && rows.length < MAX_POINTS; r++) {
			if (!this.visibleOnly || !isRowHidden(sheet, r)) {
				rows.push(r);
			}
		}
		const cols: number[] = [];
		for (let c = left; c <= right && cols.length < MAX_POINTS; c++) {
			if (!this.visibleOnly || !isColHidden(sheet, c)) {
				cols.push(c);
			}
		}
		return rows.map((r) =>
			cols.map((c) => {
				const cell = sheet.cells.get(cellKey(r, c));
				if (!cell) {
					return { value: null, format: 'General' };
				}
				return { value: cell.value, format: this.styles.resolve(cell.style).numFmt };
			}),
		);
	}

	text(cell: ReadCell): string {
		if (cell.value === null) {
			return '';
		}
		return formatValue(cell.value, cell.format, { date1904: this.workbook.date1904 }).text;
	}
}

function isRowHidden(sheet: Worksheet, row: number): boolean {
	const info = sheet.rows.get(row);
	return info?.hidden ?? Boolean(sheet.zeroHeight && !info?.customHeight);
}

function isColHidden(sheet: Worksheet, col: number): boolean {
	for (const c of sheet.cols) {
		if (col >= c.min && col <= c.max) {
			return Boolean(c.hidden);
		}
	}
	return false;
}

/** Flatten a value range (a column, a row, or a block in row-major order). */
function flatten(grid: Grid): ReadCell[] {
	if (grid.length === 0) {
		return [];
	}
	if (grid[0].length === 1) {
		return grid.map((row) => row[0]);
	}
	return grid.flat();
}

function toNumber(value: CellValue | string | number, blanks: Blanks): number | null {
	if (value === null || value === undefined) {
		return blanks === 'zero' ? 0 : blanks === 'span' ? NaN : null;
	}
	if (typeof value === 'number') {
		return Number.isFinite(value) ? value : null;
	}
	if (isCellError(value)) {
		// #N/A is interpolated over; other errors plot as zero, like Excel.
		return value.code === '#N/A' ? NaN : 0;
	}
	// Text (including "" from formulas) and booleans plot as zero.
	return 0;
}

function cacheNumbers(ref: ChartDataRef | undefined, blanks: Blanks): Array<number | null> {
	return (ref?.cache ?? []).map((v) => {
		if (typeof v === 'string') {
			const n = Number(v);
			return v.trim() !== '' && Number.isFinite(n) ? n : 0;
		}
		return toNumber(v, blanks);
	});
}

interface Numbers {
	values: Array<number | null>;
	format: string;
	live: boolean;
	hidden: boolean;
}

function readNumbers(src: CellSource, ref: ChartDataRef | undefined, blanks: Blanks): Numbers {
	const areas = src.areas(ref?.ref);
	if (!areas) {
		return {
			values: cacheNumbers(ref, blanks),
			format: ref?.formatCode ?? 'General',
			live: false,
			hidden: false,
		};
	}
	const cells = areas.flatMap((a) => flatten(src.grid(a)));
	const firstNumeric = cells.find((c) => typeof c.value === 'number');
	return {
		values: cells.map((c) => toNumber(c.value, blanks)),
		format: firstNumeric?.format ?? ref?.formatCode ?? 'General',
		live: true,
		hidden: cells.length === 0 && areas.length > 0,
	};
}

interface Categories {
	labels: string[];
	levels?: string[][];
	xValues: number[];
	numeric: boolean;
	format: string;
}

function cacheCategories(ref: ChartDataRef | undefined, date1904: boolean): Categories {
	const cache = ref?.cache ?? [];
	const format = ref?.formatCode ?? 'General';
	const labels = cache.map((v) =>
		v === null ? '' : typeof v === 'number' ? formatValue(v, format, { date1904 }).text : v,
	);
	const xValues = cache.map((v) =>
		typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN,
	);
	const out: Categories = {
		labels,
		xValues,
		numeric:
			cache.some((v) => v !== null) &&
			xValues.every((x, i) => cache[i] === null || Number.isFinite(x)),
		format,
	};
	if (ref?.levels && ref.levels.length > 1) {
		out.levels = ref.levels.slice(0, -1).map((lvl) => lvl.map((v) => v ?? ''));
	}
	return out;
}

function readCategories(src: CellSource, ref: ChartDataRef | undefined): Categories {
	const areas = src.areas(ref?.ref);
	if (!areas) {
		return cacheCategories(ref, src.workbook.date1904);
	}
	// Tall blocks hold one level per column (outer on the left); wide blocks one level per row.
	const levelCells: ReadCell[][] = [];
	for (const area of areas) {
		const grid = src.grid(area);
		const tall = area.range.bottom - area.range.top >= area.range.right - area.range.left;
		const width = grid[0]?.length ?? 0;
		const levels = tall ? width : grid.length;
		for (let l = 0; l < levels; l++) {
			const cells = tall ? grid.map((row) => row[l]) : grid[l];
			(levelCells[l] ??= []).push(...cells);
		}
	}
	const inner = levelCells[levelCells.length - 1] ?? [];
	const numericCells = inner.filter((c) => c.value !== null);
	const out: Categories = {
		labels: inner.map((c) => src.text(c)),
		xValues: inner.map((c) => (typeof c.value === 'number' ? c.value : NaN)),
		numeric: numericCells.length > 0 && numericCells.every((c) => typeof c.value === 'number'),
		format: numericCells.find((c) => typeof c.value === 'number')?.format ?? 'General',
	};
	if (levelCells.length > 1) {
		out.levels = levelCells.slice(0, -1).map((lvl) => lvl.map((c) => src.text(c)));
	}
	return out;
}

/** Text of the cells a title or series name links to (joined by spaces), if readable. */
function linkedText(src: CellSource, ref: string | undefined): string | undefined {
	const areas = src.areas(ref);
	if (!areas) {
		return undefined;
	}
	const texts = areas
		.flatMap((a) => src.grid(a).flat())
		.map((c) => src.text(c))
		.filter((t) => t !== '');
	return texts.length ? texts.join(' ') : undefined;
}

function seriesName(src: CellSource, series: ChartSeries): string {
	return linkedText(src, series.name?.ref) ?? series.name?.text ?? `Series${series.idx + 1}`;
}

function resolveSeries(src: CellSource, series: ChartSeries, blanks: Blanks): ChartSeriesData {
	const values = readNumbers(src, series.values, blanks);
	const cats = readCategories(src, series.categories);
	const data: ChartSeriesData = {
		name: seriesName(src, series),
		categories: cats.labels,
		xValues: cats.xValues,
		values: values.values,
		valueFormat: values.format,
		categoryFormat: cats.format,
		categoriesNumeric: cats.numeric,
		live: values.live,
	};
	if (cats.levels) {
		data.categoryLevels = cats.levels;
	}
	if (series.bubbleSizes) {
		data.bubbleSizes = readNumbers(src, series.bubbleSizes, 'gap').values;
	}
	if (values.hidden) {
		data.hidden = true;
	}
	return data;
}

/**
 * Resolve every series of `chart` against the current cell values of
 * `workbook`. `sheetIndex` is the sheet that holds the chart (it scopes
 * sheet-level defined names and references without a sheet name).
 */
export function resolveChartData(
	chart: ChartModel,
	workbook: Workbook,
	sheetIndex: number,
): ChartData {
	const src = new CellSource(workbook, sheetIndex, chart.plotVisibleOnly ?? true);
	const blanks: Blanks = chart.displayBlanksAs ?? 'gap';
	const data: ChartData = {
		plots: chart.plots.map((plot) => ({
			series: plot.series.map((s) => resolveSeries(src, s, blanks)),
		})),
		date1904: workbook.date1904,
	};
	const titles: NonNullable<ChartData['titles']> = {};
	const chartTitle = linkedText(src, chart.title?.ref);
	if (chartTitle !== undefined) {
		titles.chart = chartTitle;
	}
	for (const axis of chart.axes) {
		const text = linkedText(src, axis.title?.ref);
		if (text !== undefined) {
			(titles.axes ??= {})[axis.id] = text;
		}
	}
	if (titles.chart !== undefined || titles.axes) {
		data.titles = titles;
	}
	return data;
}

/** Chart data from the cached values alone (no workbook), e.g. for previews. */
export function cachedChartData(chart: ChartModel): ChartData {
	const blanks: Blanks = chart.displayBlanksAs ?? 'gap';
	return {
		plots: chart.plots.map((plot) => ({
			series: plot.series.map((s) => {
				const cats = cacheCategories(s.categories, false);
				const data: ChartSeriesData = {
					name: s.name?.text ?? `Series${s.idx + 1}`,
					categories: cats.labels,
					xValues: cats.xValues,
					values: cacheNumbers(s.values, blanks),
					valueFormat: s.values?.formatCode ?? 'General',
					categoryFormat: cats.format,
					categoriesNumeric: cats.numeric,
					live: false,
				};
				if (cats.levels) {
					data.categoryLevels = cats.levels;
				}
				if (s.bubbleSizes) {
					data.bubbleSizes = cacheNumbers(s.bubbleSizes, 'gap');
				}
				return data;
			}),
		})),
	};
}

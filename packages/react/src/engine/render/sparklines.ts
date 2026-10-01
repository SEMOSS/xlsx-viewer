import type {
	CellValue,
	ColorContext,
	SparklineGroup,
	Workbook,
	Worksheet,
} from 'xlsx-viewer-core';
import { cellKey, findSheet, parseReferenceText, resolveColorCss } from 'xlsx-viewer-core';

import type { SheetGeometry } from '../geometry/sheet-geometry';
import type { Span } from '../layout/types';

export interface SparklineBar {
	x: number;
	y: number;
	width: number;
	height: number;
	color: string;
}

export interface SparklineMarker {
	cx: number;
	cy: number;
	r: number;
	color: string;
}

/** Everything needed to draw one sparkline, in sheet pixel coordinates. */
export interface SparklineShape {
	key: string;
	/** Clip rectangle: the cell. */
	clip: { x: number; y: number; width: number; height: number };
	/** Line sparklines: one path per unbroken run. */
	paths: string[];
	lineColor: string;
	lineWidth: number;
	bars: SparklineBar[];
	markers: SparklineMarker[];
}

/** Excel's default sparkline colors (Office theme, style 1). */
const DEFAULTS = {
	series: '#376092',
	negative: '#D00000',
	marker: '#D00000',
	high: '#D00000',
	low: '#D00000',
	first: '#D00000',
	last: '#D00000',
};

/** Values of a sparkline's source range, in order (non-numbers become `null`). */
export function sparklineValues(
	workbook: Workbook,
	home: Worksheet,
	formula: string,
): Array<number | null> {
	const ref = parseReferenceText(formula);
	if (!ref) {
		return [];
	}
	const sheet = ref.sheet ? findSheet(workbook, ref.sheet.sheet) : home;
	if (!sheet) {
		return [];
	}
	const top = Math.min(ref.area.start.row, ref.area.end.row);
	const bottom = Math.max(ref.area.start.row, ref.area.end.row);
	const left = Math.min(ref.area.start.col, ref.area.end.col);
	const right = Math.max(ref.area.start.col, ref.area.end.col);
	const out: Array<number | null> = [];
	const asNumber = (v: CellValue | undefined): number | null =>
		typeof v === 'number' && Number.isFinite(v) ? v : null;
	// A row of cells reads left to right, a column top to bottom.
	if (bottom - top >= right - left) {
		for (let r = top; r <= Math.min(bottom, top + 10_000); r++) {
			out.push(asNumber(sheet.cells.get(cellKey(r, left))?.value));
		}
	} else {
		for (let c = left; c <= Math.min(right, left + 10_000); c++) {
			out.push(asNumber(sheet.cells.get(cellKey(top, c))?.value));
		}
	}
	return out;
}

function axisRange(
	group: SparklineGroup,
	values: Array<number | null>,
	groupValues: Array<number | null>,
): { min: number; max: number } {
	const nums = (list: Array<number | null>) => list.filter((v): v is number => v !== null);
	const own = nums(values);
	const all =
		group.minAxisType === 'group' || group.maxAxisType === 'group' ? nums(groupValues) : own;
	let min =
		group.minAxisType === 'custom' && group.manualMin !== undefined
			? group.manualMin
			: Math.min(...(group.minAxisType === 'group' ? all : own));
	let max =
		group.maxAxisType === 'custom' && group.manualMax !== undefined
			? group.manualMax
			: Math.max(...(group.maxAxisType === 'group' ? all : own));
	if (!Number.isFinite(min) || !Number.isFinite(max)) {
		min = 0;
		max = 1;
	}
	if (group.type === 'column' && group.minAxisType !== 'custom') {
		// Columns grow from zero unless the data goes below it.
		min = Math.min(0, min);
	}
	if (max === min) {
		max = min + 1;
	}
	return { min, max };
}

/**
 * Geometry of the sparklines whose cells fall in the visible span. Values
 * come from the (calculated) cells of each sparkline's source range.
 */
export function layoutSparklines(
	workbook: Workbook,
	sheet: Worksheet,
	geometry: SheetGeometry,
	rows: Span,
	cols: Span,
	ctx: ColorContext,
): SparklineShape[] {
	const shapes: SparklineShape[] = [];
	const zoom = geometry.zoom;
	sheet.sparklineGroups.forEach((group, gi) => {
		const visible = group.sparklines.filter(
			(s) =>
				s.location.row >= rows.start &&
				s.location.row <= rows.end &&
				s.location.col >= cols.start &&
				s.location.col <= cols.end,
		);
		if (visible.length === 0) {
			return;
		}
		const color = (spec: SparklineGroup['seriesColor'], fallback: string) =>
			resolveColorCss(spec, ctx, fallback) ?? fallback;
		const series = color(group.seriesColor, DEFAULTS.series);
		const groupValues =
			group.minAxisType === 'group' || group.maxAxisType === 'group'
				? group.sparklines.flatMap((s) => sparklineValues(workbook, sheet, s.formula))
				: [];
		for (const s of visible) {
			const raw = sparklineValues(workbook, sheet, s.formula);
			const values = group.rightToLeft ? [...raw].reverse() : raw;
			const cell = geometry.cellRect(s.location.row, s.location.col);
			const shape: SparklineShape = {
				key: `${gi}:${s.location.row}:${s.location.col}`,
				clip: cell,
				paths: [],
				lineColor: series,
				lineWidth: Math.max(1, ((group.lineWeight ?? 0.75) * 96) / 72) * zoom,
				bars: [],
				markers: [],
			};
			const n = values.length;
			const nums = values.filter((v): v is number => v !== null);
			if (n === 0 || nums.length === 0) {
				shapes.push(shape);
				continue;
			}
			const padX = Math.max(2, 3 * zoom);
			const padY = Math.max(2, 3 * zoom);
			const left = cell.x + padX;
			const width = Math.max(1, cell.width - padX * 2);
			const top = cell.y + padY;
			const height = Math.max(1, cell.height - padY * 2);
			const { min, max } = axisRange(group, values, groupValues);
			const yOf = (v: number) =>
				top + height - ((Math.max(min, Math.min(max, v)) - min) / (max - min)) * height;
			const high = Math.max(...nums);
			const low = Math.min(...nums);
			const firstIndex = values.findIndex((v) => v !== null);
			const lastIndex = values.length - 1 - [...values].reverse().findIndex((v) => v !== null);
			const pointColor = (v: number, i: number): string | undefined => {
				if (group.high && v === high) return color(group.highColor, DEFAULTS.high);
				if (group.low && v === low) return color(group.lowColor, DEFAULTS.low);
				if (group.first && i === firstIndex) return color(group.firstColor, DEFAULTS.first);
				if (group.last && i === lastIndex) return color(group.lastColor, DEFAULTS.last);
				if (group.negative && v < 0) return color(group.negativeColor, DEFAULTS.negative);
				return undefined;
			};
			if (group.type === 'line') {
				const step = n > 1 ? width / (n - 1) : 0;
				const xOf = (i: number) => (n > 1 ? left + i * step : left + width / 2);
				let d = '';
				values.forEach((raw, i) => {
					const v = raw ?? (group.displayEmptyCellsAs === 'zero' ? 0 : null);
					if (v === null) {
						if (group.displayEmptyCellsAs !== 'span' && d) {
							shape.paths.push(d);
							d = '';
						}
						return;
					}
					d += `${d ? 'L' : 'M'}${xOf(i).toFixed(2)} ${yOf(v).toFixed(2)}`;
					const special = pointColor(v, i);
					if (special || group.markers) {
						shape.markers.push({
							cx: xOf(i),
							cy: yOf(v),
							r: Math.max(1.5, 2 * zoom),
							color: special ?? color(group.markersColor, DEFAULTS.marker),
						});
					}
				});
				if (d) {
					shape.paths.push(d);
				}
			} else {
				const slot = width / n;
				const barW = Math.max(1, slot * 0.8);
				const mid = top + height / 2;
				const zeroY = group.type === 'column' ? yOf(Math.max(min, Math.min(max, 0))) : mid;
				values.forEach((v, i) => {
					if (v === null || (group.type === 'stacked' && v === 0)) {
						return;
					}
					const x = left + i * slot + (slot - barW) / 2;
					const fill = pointColor(v, i) ?? series;
					if (group.type === 'stacked') {
						// Win/loss: equal bars above (win) or below (loss) the middle.
						shape.bars.push({
							x,
							y: v > 0 ? top : mid,
							width: barW,
							height: height / 2,
							color: fill,
						});
					} else {
						const y = yOf(v);
						shape.bars.push({
							x,
							y: Math.min(y, zeroY),
							width: barW,
							height: Math.max(1, Math.abs(zeroY - y)),
							color: fill,
						});
					}
				});
			}
			shapes.push(shape);
		}
	});
	return shapes;
}

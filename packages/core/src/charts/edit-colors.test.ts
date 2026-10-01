import { describe, expect, it } from 'vitest';
import { resolveDrawingColorCss } from '../color/resolve-color';
import { createDefaultTheme } from '../styles/defaults';
import type { ChartModel } from '../types/chart';
import type { ShapeFill } from '../types/drawing';
import { byRole, loadFixture, makeWorkbook } from './__tests__/helpers';
import { createChartFromRange } from './data/create-chart';
import { cachedChartData } from './data/resolve-data';
import {
	applyChartStyle,
	applyColorPalette,
	CHART_COLOR_PALETTES,
	chartPaletteColor,
	convertChartType,
} from './edit';
import { layoutChart } from './layout/layout-chart';
import type { ChartRectNode } from './layout/types';
import { parseChartXml } from './parse/parse-chart';
import { writeChartXml } from './write/write-chart';

const theme = createDefaultTheme();
const wb = makeWorkbook({
	Sheet1: [
		[null, 'A', 'B', 'C', 'D'],
		['Q1', 1, 2, 3, 4],
		['Q2', 2, 3, 4, 5],
		['Q3', 3, 4, 5, 6],
		['Q4', 4, 5, 6, 7],
		['Q5', 5, 6, 7, 8],
	],
});
const chart = (): ChartModel =>
	createChartFromRange(wb, 0, { top: 0, left: 0, bottom: 5, right: 4 }, 'column');
const css = (fill: ShapeFill | undefined) =>
	fill?.type === 'solid' ? resolveDrawingColorCss(fill.color, theme) : undefined;

describe('chartPaletteColor', () => {
	it('follows the colorful and monochromatic palettes', () => {
		const cycle = (palette: (typeof CHART_COLOR_PALETTES)[number]) =>
			[0, 1, 2].map((i) => resolveDrawingColorCss(chartPaletteColor(palette, i, 3), theme));
		expect(cycle('colorful1')).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
		expect(cycle('colorful2')).toEqual(['#4472C4', '#A5A5A5', '#5B9BD5']);
		expect(cycle('colorful3')).toEqual(['#ED7D31', '#FFC000', '#70AD47']);
		expect(cycle('colorful4')).toEqual(['#70AD47', '#5B9BD5', '#FFC000']);
		expect(chartPaletteColor('colorful2', 3, 6)).toEqual({
			kind: 'scheme',
			value: 'accent1',
			transforms: [{ name: 'lumMod', val: 60000 }],
		});
		const mono = cycle('monochromatic2');
		expect(new Set(mono).size).toBe(3);
		expect(chartPaletteColor('monochromatic2', 0, 3)).toMatchObject({ value: 'accent2' });
		expect(CHART_COLOR_PALETTES).toHaveLength(10);
	});
});

describe('applyColorPalette', () => {
	it('writes explicit series fills without touching the input', () => {
		const source = chart();
		const before = JSON.stringify(source);
		const result = applyColorPalette(source, 'colorful3');
		expect(JSON.stringify(source)).toBe(before);
		expect(result.dirty).toBe(true);
		expect(result.plots[0].series.map((s) => css(s.fill))).toEqual([
			'#ED7D31',
			'#FFC000',
			'#70AD47',
			'#9E480E',
		]);
		const bars = byRole(
			layoutChart(result, cachedChartData(result), theme, 480, 288),
			'bar',
		) as ChartRectNode[];
		expect(bars[0].fill).toBe('#ED7D31');
		expect(parseChartXml(writeChartXml(result)).plots[0].series[2].fill).toEqual(
			result.plots[0].series[2].fill,
		);
	});

	it('colors every point of vary-colors charts', () => {
		const pie = applyColorPalette(convertChartType(chart(), 'pie'), 'monochromatic1');
		const points = pie.plots[0].series[0].points ?? [];
		expect(points.map((p) => p.idx)).toEqual([0, 1, 2, 3, 4]);
		expect(new Set(points.map((p) => css(p.fill))).size).toBe(5);
		const office = applyColorPalette(parseChartXml(loadFixture('pie-office.xml')), 'colorful2');
		const kept = office.plots[0].series[0].points?.[0];
		expect(kept?.line?.width).toBe(19050);
		expect(css(kept?.fill)).toBe('#4472C4');
	});

	it('recolors lines and markers but keeps hidden scatter lines hidden', () => {
		const lines = applyColorPalette(convertChartType(chart(), 'lineMarkers'), 'colorful4').plots[0]
			.series[0];
		expect(css(lines.line?.fill)).toBe('#70AD47');
		expect(css(lines.marker?.fill)).toBe('#70AD47');
		const scatter = applyColorPalette(convertChartType(chart(), 'scatter'), 'colorful1').plots[0]
			.series[1];
		expect(scatter.line?.fill).toEqual({ type: 'none' });
		expect(css(scatter.marker?.fill)).toBe('#ED7D31');
	});
});

describe('applyChartStyle', () => {
	it('applies quick styles over earlier ones', () => {
		const source = chart();
		const before = JSON.stringify(source);
		const dark = applyChartStyle(source, 6);
		expect(JSON.stringify(source)).toBe(before);
		expect(css(dark.chartFill)).toBe('#404040');
		expect(resolveDrawingColorCss(dark.axes[0].textStyle?.color, theme)).toBe('#D9D9D9');
		const labelled = applyChartStyle(dark, 2);
		expect(css(labelled.chartFill)).toBe('#FFFFFF');
		expect(labelled.axes[1].majorGridlines).toBeUndefined();
		expect(labelled.plots[0].series.every((s) => s.dataLabels?.showValue)).toBe(true);
		expect(labelled.plots[0].gapWidth).toBe(100);
		const plain = applyChartStyle(labelled, 1);
		expect(plain.plots[0].series[0].dataLabels).toBeUndefined();
		expect(plain.axes[1].majorGridlines).toBeDefined();
		expect(plain.plots[0].gapWidth).toBe(219);
		expect(applyChartStyle(source, 4).plots[0].series[0].line?.width).toBe(19050);
		expect(applyChartStyle(source, 3).axes[1].majorGridlines).toMatchObject({ dash: 'dash' });
		expect(applyChartStyle(source, 99).chartFill).toEqual(dark.chartFill);
	});

	it('labels pies with percentages and produces valid parts', () => {
		const pie = applyChartStyle(convertChartType(chart(), 'pie'), 2);
		expect(pie.plots[0].series[0].dataLabels).toMatchObject({
			showPercent: true,
			showValue: false,
		});
		const back = parseChartXml(writeChartXml(applyChartStyle(chart(), 5)));
		expect(css(back.chartFill)).toBe('#F2F2F2');
	});
});

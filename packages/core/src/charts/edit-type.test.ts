import { describe, expect, it } from 'vitest';
import { createDefaultTheme } from '../styles/defaults';
import type { ChartModel } from '../types/chart';
import { assertFinite, loadFixture, makeWorkbook } from './__tests__/helpers';
import { createChartFromRange } from './data/create-chart';
import { cachedChartData } from './data/resolve-data';
import { CHART_TYPE_TARGETS, chartTypeTargetOf, convertChartType } from './edit';
import { layoutChart } from './layout/layout-chart';
import { parseChartXml } from './parse/parse-chart';
import { writeChartXml } from './write/write-chart';

const wb = makeWorkbook({
	Sheet1: [
		[null, 'Revenue', 'Cost'],
		['Q1', 10, 4],
		['Q2', 20, 8],
		['Q3', 30, 12],
	],
});
const column = () =>
	createChartFromRange(wb, 0, { top: 0, left: 0, bottom: 3, right: 2 }, 'column', {
		title: 'Sales',
	});

describe('convertChartType', () => {
	it.each(CHART_TYPE_TARGETS)('converts a column chart to %s', (target) => {
		const source = column();
		const before = JSON.stringify(source);
		const result = convertChartType(source, target);
		expect(JSON.stringify(source)).toBe(before);
		expect(result.dirty).toBe(true);
		expect(result.plots).toHaveLength(1);
		expect(chartTypeTargetOf(result)).toBe(target);
		const plot = result.plots[0];
		expect(plot.series.map((s) => s.name?.text)).toEqual(['Revenue', 'Cost']);
		expect(plot.series[1].values?.ref).toBe('Sheet1!$C$2:$C$4');
		expect(result.title?.text).toBe('Sales');
		const kinds = result.axes.map((a) => a.kind).join(',');
		expect(kinds).toBe(
			target === 'pie' || target === 'doughnut'
				? ''
				: target.startsWith('scatter')
					? 'val,val'
					: 'cat,val',
		);
		expect(plot.axisIds).toEqual(result.axes.map((a) => a.id));
		const back = parseChartXml(writeChartXml(result));
		expect(chartTypeTargetOf(back)).toBe(target);
		assertFinite(layoutChart(result, cachedChartData(result), createDefaultTheme(), 400, 260));
	});

	it('carries series colors over to lines and markers', () => {
		const line = convertChartType(column(), 'lineMarkers').plots[0].series[0];
		expect(line.line?.fill).toEqual({ type: 'solid', color: { kind: 'scheme', value: 'accent1' } });
		expect(line.marker).toMatchObject({
			symbol: 'circle',
			fill: { type: 'solid', color: { kind: 'scheme', value: 'accent1' } },
		});
		const plain = convertChartType(column(), 'line').plots[0].series[1];
		expect(plain.marker).toEqual({ symbol: 'none' });
		const back = convertChartType(convertChartType(column(), 'line'), 'columnStacked').plots[0];
		expect(back.series[1].fill).toEqual({
			type: 'solid',
			color: { kind: 'scheme', value: 'accent2' },
		});
		expect(back).toMatchObject({ grouping: 'stacked', overlap: 100 });
	});

	it('lets pies vary colors by point instead of using series fills', () => {
		const pie = convertChartType(column(), 'pie').plots[0];
		expect(pie).toMatchObject({ type: 'pie', varyColors: true });
		expect(pie.series[0].fill).toBeUndefined();
		const fromPie = convertChartType(parseChartXml(loadFixture('pie-office.xml')), 'barClustered');
		expect(fromPie.plots[0].varyColors).toBe(false);
		expect(fromPie.plots[0].series[0].points).toBeUndefined();
		expect(fromPie.axes.map((a) => `${a.kind}@${a.position}`)).toEqual(['cat@l', 'val@b']);
	});

	it('merges combo plots and drops secondary and 3-D axes', () => {
		const combo = convertChartType(parseChartXml(loadFixture('combo-bar-line.xml')), 'area');
		expect(combo.plots).toHaveLength(1);
		expect(combo.plots[0].series.map((s) => s.idx)).toEqual([0, 1, 2]);
		expect(combo.axes[1].crossBetween).toBe('midCat');
		const flat = convertChartType(
			parseChartXml(loadFixture('bar3d-horizontal.xml')),
			'columnClustered',
		);
		expect(flat.plots[0].is3D).toBeUndefined();
		expect(flat.axes.map((a) => a.kind)).toEqual(['cat', 'val']);
		expect(flat.axes.map((a) => a.position)).toEqual(['b', 'l']);
	});

	it('turns the X value axis of a scatter chart into a category axis and back', () => {
		const scatter = parseChartXml(loadFixture('scatter.xml'));
		scatter.axes[0].crossBetween = 'midCat';
		scatter.axes[0].min = 1;
		const column = convertChartType(scatter, 'columnClustered');
		expect(column.axes[0]).toMatchObject({ kind: 'cat', id: scatter.axes[0].id });
		expect(column.axes[0].min).toBeUndefined();
		expect(column.axes[1].crossBetween).toBe('between');
		const again = convertChartType(column, 'scatterLines');
		expect(again.axes.map((a) => a.kind)).toEqual(['val', 'val']);
		expect(again.plots[0].series[0].line?.fill?.type).toBe('solid');
	});

	it('removes data label positions the new type does not allow', () => {
		const source = column();
		source.plots[0].series[0].dataLabels = { showValue: true, position: 'outEnd' };
		expect(convertChartType(source, 'columnStacked').plots[0].series[0].dataLabels).toEqual({
			showValue: true,
		});
		expect(convertChartType(source, 'barClustered').plots[0].series[0].dataLabels?.position).toBe(
			'outEnd',
		);
	});

	it('keeps trendlines only where Excel allows them', () => {
		const source = column();
		source.plots[0].series[0].trendlines = [{ type: 'linear' }];
		expect(convertChartType(source, 'line').plots[0].series[0].trendlines).toHaveLength(1);
		expect(convertChartType(source, 'columnStacked').plots[0].series[0].trendlines).toBeUndefined();
		expect(convertChartType(source, 'pie').plots[0].series[0].trendlines).toBeUndefined();
	});

	it('creates default axes for charts that had none', () => {
		const pie: ChartModel = convertChartType(column(), 'doughnut');
		const bars = convertChartType(pie, 'columnPercent');
		expect(bars.axes.map((a) => `${a.kind}@${a.position}`)).toEqual(['cat@b', 'val@l']);
		expect(bars.axes[0].crossAxisId).toBe(bars.axes[1].id);
	});
});

describe('chartTypeTargetOf', () => {
	it('recognises the fixtures', () => {
		const types = [
			'area.xml',
			'bar3d-horizontal.xml',
			'line-stacked-markers.xml',
			'scatter.xml',
			'stock.xml',
			'doughnut-office.xml',
		].map((name) => chartTypeTargetOf(parseChartXml(loadFixture(name))));
		expect(types).toEqual([
			'area',
			'barClustered',
			'lineMarkers',
			'scatter',
			undefined,
			'doughnut',
		]);
		expect(chartTypeTargetOf({ plots: [], axes: [] })).toBeUndefined();
	});
});

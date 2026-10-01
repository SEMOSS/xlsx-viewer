import { describe, expect, it } from 'vitest';
import { createDefaultTheme } from '../../styles/defaults';
import { assertFinite, byRole, makeWorkbook, texts } from '../__tests__/helpers';
import { layoutChart } from '../layout/layout-chart';
import { parseChartXml } from '../parse/parse-chart';
import { writeChartXml } from '../write/write-chart';
import { createChartFromRange } from './create-chart';
import { resolveChartData } from './resolve-data';

const sales = makeWorkbook({
	'Sales Data': [
		[null, 'Revenue', 'Cost'],
		['Q1', 10, 4],
		['Q2', 20, 8],
		['Q3', 30, 12],
		['Q4', 40, 16],
	],
});
const all = { top: 0, left: 0, bottom: 4, right: 2 };

describe('createChartFromRange', () => {
	it('uses the header row for series names and the first column for categories', () => {
		const chart = createChartFromRange(sales, 0, all, 'column');
		expect(chart.dirty).toBe(true);
		expect(chart.rawXml).toBeUndefined();
		const plot = chart.plots[0];
		expect(plot).toMatchObject({
			type: 'bar',
			barDirection: 'col',
			grouping: 'clustered',
			gapWidth: 219,
			overlap: -27,
		});
		expect(plot.series.map((s) => s.name)).toEqual([
			{ ref: "'Sales Data'!$B$1", text: 'Revenue' },
			{ ref: "'Sales Data'!$C$1", text: 'Cost' },
		]);
		expect(plot.series[0].categories).toEqual({
			kind: 'str',
			ref: "'Sales Data'!$A$2:$A$5",
			cache: ['Q1', 'Q2', 'Q3', 'Q4'],
		});
		expect(plot.series[1].values).toEqual({
			kind: 'num',
			ref: "'Sales Data'!$C$2:$C$5",
			cache: [4, 8, 12, 16],
			formatCode: 'General',
		});
		expect(chart.axes.map((a) => `${a.kind}@${a.position}`)).toEqual(['cat@b', 'val@l']);
		expect(plot.axisIds).toEqual(chart.axes.map((a) => a.id));
		expect(chart.legend?.position).toBe('b');
		expect(chart.title).toMatchObject({ overlay: false });
		expect(chart.title?.text).toBeUndefined();
	});

	it('expands a single cell to the current region', () => {
		const chart = createChartFromRange(sales, 0, { top: 2, left: 1, bottom: 2, right: 1 }, 'line');
		expect(chart.plots[0].series).toHaveLength(2);
		expect(chart.plots[0].series[0].values?.ref).toBe("'Sales Data'!$B$2:$B$5");
		expect(chart.plots[0]).toMatchObject({ type: 'line', grouping: 'standard', marker: true });
		expect(chart.plots[0].series[0].marker).toEqual({ symbol: 'none' });
	});

	it('plots series in rows when the data is wider than tall', () => {
		const wide = makeWorkbook({
			Sheet1: [
				[null, 2021, 2022, 2023, 2024],
				['North', 1, 2, 3, 4],
				['South', 5, 6, 7, 8],
			],
		});
		const chart = createChartFromRange(wide, 0, { top: 0, left: 0, bottom: 2, right: 4 }, 'column');
		const [north, south] = chart.plots[0].series;
		expect(chart.plots[0].series).toHaveLength(2);
		expect(north.name).toEqual({ ref: 'Sheet1!$A$2', text: 'North' });
		expect(north.values?.ref).toBe('Sheet1!$B$2:$E$2');
		expect(south.values?.cache).toEqual([5, 6, 7, 8]);
		expect(north.categories).toMatchObject({
			kind: 'num',
			ref: 'Sheet1!$B$1:$E$1',
			cache: [2021, 2022, 2023, 2024],
		});
		const columns = createChartFromRange(
			wide,
			0,
			{ top: 0, left: 0, bottom: 2, right: 4 },
			'column',
			{ seriesIn: 'columns' },
		);
		expect(columns.plots[0].series).toHaveLength(4);
		expect(columns.plots[0].series[0].name?.text).toBe('2021');
	});

	it('builds horizontal bars, pies, doughnuts, areas, radars and scatter charts', () => {
		const bar = createChartFromRange(sales, 0, all, 'bar', { grouping: 'stacked' });
		expect(bar.plots[0]).toMatchObject({
			barDirection: 'bar',
			grouping: 'stacked',
			overlap: 100,
			gapWidth: 150,
		});
		expect(bar.axes.map((a) => a.position)).toEqual(['l', 'b']);
		const pie = createChartFromRange(sales, 0, all, 'pie');
		expect(pie.plots[0]).toMatchObject({ type: 'pie', varyColors: true, firstSliceAngle: 0 });
		expect(pie.axes).toEqual([]);
		expect(pie.legend?.position).toBe('b');
		expect(createChartFromRange(sales, 0, all, 'doughnut').plots[0]).toMatchObject({
			type: 'doughnut',
			holeSize: 75,
		});
		const area = createChartFromRange(sales, 0, all, 'area');
		expect(area.axes[1].crossBetween).toBe('midCat');
		const radar = createChartFromRange(sales, 0, all, 'radar');
		expect(radar.plots[0].radarStyle).toBe('marker');
		const xy = makeWorkbook({
			Sheet1: [
				['x', 'y'],
				[1, 2],
				[2, 4],
				[3, 9],
			],
		});
		const scatter = createChartFromRange(
			xy,
			0,
			{ top: 0, left: 0, bottom: 3, right: 1 },
			'scatter',
		);
		expect(scatter.plots[0].series).toHaveLength(1);
		expect(scatter.plots[0].series[0].categories).toMatchObject({
			kind: 'num',
			ref: 'Sheet1!$A$2:$A$4',
			cache: [1, 2, 3],
		});
		expect(scatter.plots[0].series[0].name?.text).toBe('y');
		expect(scatter.axes.map((a) => a.kind)).toEqual(['val', 'val']);
		expect(scatter.plots[0].series[0].line?.fill).toEqual({ type: 'none' });
	});

	it('honours title and legend options and omits the legend for one series', () => {
		const one = createChartFromRange(sales, 0, { top: 0, left: 0, bottom: 4, right: 1 }, 'column');
		expect(one.legend).toBeUndefined();
		const titled = createChartFromRange(sales, 0, all, 'column', {
			title: 'Quarterly',
			legend: 'r',
		});
		expect(titled.title?.text).toBe('Quarterly');
		expect(titled.legend?.position).toBe('r');
		const bare = createChartFromRange(sales, 0, all, 'column', { title: false, legend: false });
		expect(bare.title).toBeUndefined();
		expect(bare.autoTitleDeleted).toBe(true);
		expect(bare.legend).toBeUndefined();
	});

	it('uses several leading label columns as multi-level categories and dates as categories', () => {
		const book = makeWorkbook({
			Sheet1: [
				['Year', 'Month', 'Sales'],
				['2023', 'Jan', 1],
				[null, 'Feb', 2],
				['2024', 'Jan', 3],
			],
			Dates: [
				['Day', 'Visits'],
				[{ v: 45292, fmt: 'm/d/yyyy' }, 10],
				[{ v: 45293, fmt: 'm/d/yyyy' }, 12],
			],
		});
		const multi = createChartFromRange(book, 0, { top: 0, left: 0, bottom: 3, right: 2 }, 'column');
		expect(multi.plots[0].series).toHaveLength(1);
		expect(multi.plots[0].series[0].categories).toEqual({
			kind: 'multiLvlStr',
			ref: 'Sheet1!$A$2:$B$4',
			cache: ['Jan', 'Feb', 'Jan'],
			levels: [
				['2023', null, '2024'],
				['Jan', 'Feb', 'Jan'],
			],
		});
		const dated = createChartFromRange(book, 1, { top: 0, left: 0, bottom: 2, right: 1 }, 'line');
		expect(dated.plots[0].series[0].categories).toEqual({
			kind: 'num',
			ref: 'Dates!$A$2:$A$3',
			cache: [45292, 45293],
			formatCode: 'm/d/yyyy',
		});
	});

	it('writes a valid chart part that parses back to the same model', () => {
		for (const type of [
			'column',
			'bar',
			'line',
			'pie',
			'doughnut',
			'area',
			'scatter',
			'radar',
		] as const) {
			const chart = createChartFromRange(sales, 0, all, type);
			const xml = writeChartXml(chart);
			expect(xml).toContain('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"');
			const back = parseChartXml(xml);
			const { rawXml, dirty, ...rest } = back;
			expect(rawXml).toBe(xml);
			expect(dirty).toBe(false);
			const { dirty: _d, ...expected } = chart;
			expect(rest).toEqual(expected);
		}
	});

	it('renders the created chart with live data', () => {
		const chart = createChartFromRange(sales, 0, all, 'column');
		const tree = layoutChart(
			chart,
			resolveChartData(chart, sales, 0),
			createDefaultTheme(),
			480,
			288,
		);
		assertFinite(tree);
		expect(byRole(tree, 'bar')).toHaveLength(8);
		expect(texts(tree, 'legend-entry')).toEqual(['Revenue', 'Cost']);
		expect(texts(tree, 'title')).toEqual(['Chart Title']);
	});

	it('rejects unknown sheets', () => {
		expect(() => createChartFromRange(sales, 5, all, 'column')).toThrow(RangeError);
	});
});

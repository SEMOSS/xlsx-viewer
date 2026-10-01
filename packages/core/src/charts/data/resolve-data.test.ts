import { describe, expect, it } from 'vitest';
import { ERRORS } from '../../values/errors';
import type { ChartModel, ChartSeries } from '../../types/chart';
import { loadFixture, makeWorkbook } from '../__tests__/helpers';
import { parseChartXml } from '../parse/parse-chart';
import { cachedChartData, resolveChartData } from './resolve-data';

function chartOf(series: ChartSeries[], extra: Partial<ChartModel> = {}): ChartModel {
	return { plots: [{ type: 'line', series, axisIds: [] }], axes: [], ...extra };
}

const lineSeries = (values: string, categories?: string, name?: string): ChartSeries => ({
	idx: 0,
	order: 0,
	name: name ? { ref: name, text: 'cached name' } : undefined,
	categories: categories
		? { kind: 'str', ref: categories, cache: ['c1', 'c2', 'c3', 'c4'] }
		: undefined,
	values: { kind: 'num', ref: values, cache: [9, 9, 9, 9], formatCode: 'General' },
});

describe('resolveChartData', () => {
	const wb = makeWorkbook({
		Sheet1: [
			['', 'Revenue', 'Cost'],
			['Q1', 10, 1],
			['Q2', 20, 2],
			['Q3', 30, 3],
			['Q4', 40, 4],
		],
	});

	it('reads live cell values, names and categories', () => {
		const data = resolveChartData(
			chartOf([lineSeries('Sheet1!$B$2:$B$5', 'Sheet1!$A$2:$A$5', 'Sheet1!$B$1')]),
			wb,
			0,
		);
		expect(data.plots[0].series[0]).toMatchObject({
			name: 'Revenue',
			categories: ['Q1', 'Q2', 'Q3', 'Q4'],
			values: [10, 20, 30, 40],
			live: true,
			categoriesNumeric: false,
			valueFormat: 'General',
		});
	});

	it('resolves the series of a parsed chart and falls back to its cache for unknown sheets', () => {
		const chart = parseChartXml(loadFixture('area.xml'));
		const live = resolveChartData(chart, wb, 0).plots[0].series[0];
		expect(live.values).toEqual([10, 20, 30, 40]);
		const other = makeWorkbook({ Data: [[1]] });
		const cached = resolveChartData(chart, other, 0).plots[0].series[0];
		expect(cached).toMatchObject({
			name: 'Revenue',
			values: [45, 62, 58, 71],
			categories: ['Q1', 'Q2', 'Q3', 'Q4'],
			live: false,
		});
	});

	it('formats numeric and date categories with the cell number format', () => {
		const dated = makeWorkbook({
			Sheet1: [
				[
					{ v: 45292, fmt: 'm/d/yyyy' },
					{ v: 1234.5, fmt: '#,##0.00' },
				],
				[{ v: 45293, fmt: 'm/d/yyyy' }, 2],
			],
		});
		const s = resolveChartData(
			chartOf([lineSeries('Sheet1!$B$1:$B$2', 'Sheet1!$A$1:$A$2')]),
			dated,
			0,
		).plots[0].series[0];
		expect(s.categories).toEqual(['1/1/2024', '1/2/2024']);
		expect(s.categoriesNumeric).toBe(true);
		expect(s.categoryFormat).toBe('m/d/yyyy');
		expect(s.xValues).toEqual([45292, 45293]);
		expect(s.valueFormat).toBe('#,##0.00');
	});

	it('drops hidden rows unless the chart plots hidden cells', () => {
		const book = makeWorkbook({
			Sheet1: [
				['A', 1],
				['B', 2],
				['C', 3],
			],
		});
		book.sheets[0].rows.set(1, { hidden: true });
		const series = lineSeries('Sheet1!$B$1:$B$3', 'Sheet1!$A$1:$A$3');
		const visible = resolveChartData(chartOf([series]), book, 0).plots[0].series[0];
		expect(visible.values).toEqual([1, 3]);
		expect(visible.categories).toEqual(['A', 'C']);
		const all = resolveChartData(chartOf([series], { plotVisibleOnly: false }), book, 0).plots[0]
			.series[0];
		expect(all.values).toEqual([1, 2, 3]);
	});

	it('hides a series whose source column is hidden', () => {
		const book = makeWorkbook({
			Sheet1: [
				['A', 1],
				['B', 2],
			],
		});
		book.sheets[0].cols.push({ min: 1, max: 1, hidden: true });
		const s = resolveChartData(
			chartOf([lineSeries('Sheet1!$B$1:$B$2', 'Sheet1!$A$1:$A$2')]),
			book,
			0,
		).plots[0].series[0];
		expect(s.hidden).toBe(true);
		expect(s.values).toEqual([]);
	});

	it('applies displayBlanksAs and Excel rules for text and errors', () => {
		const book = makeWorkbook({ Sheet1: [[1, null, 'text', ERRORS.NA, ERRORS.DIV0, '', true]] });
		const series = lineSeries('Sheet1!$A$1:$G$1');
		const values = (blanks?: ChartModel['displayBlanksAs']) =>
			resolveChartData(chartOf([series], { displayBlanksAs: blanks }), book, 0).plots[0].series[0]
				.values;
		expect(values()).toEqual([1, null, 0, NaN, 0, 0, 0]);
		expect(values('zero')).toEqual([1, 0, 0, NaN, 0, 0, 0]);
		expect(values('span')).toEqual([1, NaN, 0, NaN, 0, 0, 0]);
	});

	it('reads multi-area references and defined names', () => {
		const book = makeWorkbook({ Sheet1: [[1], [2], [3], [4]] });
		book.definedNames.push({ name: 'Picks', formula: '(Sheet1!$A$1,Sheet1!$A$4)' });
		expect(
			resolveChartData(chartOf([lineSeries('(Sheet1!$A$1,Sheet1!$A$3:$A$4)')]), book, 0).plots[0]
				.series[0].values,
		).toEqual([1, 3, 4]);
		expect(
			resolveChartData(chartOf([lineSeries('Sheet1!Picks')]), book, 0).plots[0].series[0].values,
		).toEqual([1, 4]);
	});

	it('limits whole-column references to the used range', () => {
		const book = makeWorkbook({ Sheet1: [[5], [6], [7]] });
		expect(
			resolveChartData(chartOf([lineSeries('Sheet1!$A:$A')]), book, 0).plots[0].series[0].values,
		).toEqual([5, 6, 7]);
	});

	it('builds multi-level categories from a block (outer levels first)', () => {
		const book = makeWorkbook({
			Sheet1: [
				['2023', 'Jan', 1],
				[null, 'Feb', 2],
				['2024', 'Jan', 3],
			],
		});
		const s = resolveChartData(
			chartOf([lineSeries('Sheet1!$C$1:$C$3', 'Sheet1!$A$1:$B$3')]),
			book,
			0,
		).plots[0].series[0];
		expect(s.categories).toEqual(['Jan', 'Feb', 'Jan']);
		expect(s.categoryLevels).toEqual([['2023', '', '2024']]);
	});

	it('joins multi-cell series names and defaults unnamed series', () => {
		const book = makeWorkbook({ Sheet1: [['North', 'Sales', 1]] });
		const named = lineSeries('Sheet1!$C$1', undefined, 'Sheet1!$A$1:$B$1');
		expect(resolveChartData(chartOf([named]), book, 0).plots[0].series[0].name).toBe('North Sales');
		const unnamed: ChartSeries = {
			idx: 2,
			order: 0,
			values: { kind: 'num', ref: 'Sheet1!$C$1', cache: [] },
		};
		expect(resolveChartData(chartOf([unnamed]), book, 0).plots[0].series[0].name).toBe('Series3');
	});

	it('reads cell-linked chart and axis titles live', () => {
		const book = makeWorkbook({ Sheet1: [['Live title', 'Units', 1]] });
		const chart = chartOf([lineSeries('Sheet1!$C$1')], {
			title: { ref: 'Sheet1!$A$1', text: 'Cached title' },
			axes: [
				{
					id: 2,
					kind: 'val',
					position: 'l',
					crossAxisId: 1,
					title: { ref: 'Sheet1!$B$1', text: 'Old' },
				},
			],
		});
		expect(resolveChartData(chart, book, 0).titles).toEqual({
			chart: 'Live title',
			axes: { 2: 'Units' },
		});
		expect(resolveChartData(chartOf([lineSeries('Sheet1!$C$1')]), book, 0).titles).toBeUndefined();
	});

	it('reads bubble sizes', () => {
		const book = makeWorkbook({
			Sheet1: [
				[1, 10, 5],
				[2, 20, 8],
			],
		});
		const s: ChartSeries = {
			idx: 0,
			order: 0,
			categories: { kind: 'num', ref: 'Sheet1!$A$1:$A$2', cache: [] },
			values: { kind: 'num', ref: 'Sheet1!$B$1:$B$2', cache: [] },
			bubbleSizes: { kind: 'num', ref: 'Sheet1!$C$1:$C$2', cache: [] },
		};
		const data = resolveChartData(
			{ plots: [{ type: 'bubble', series: [s], axisIds: [] }], axes: [] },
			book,
			0,
		).plots[0].series[0];
		expect(data.xValues).toEqual([1, 2]);
		expect(data.bubbleSizes).toEqual([5, 8]);
	});
});

describe('cachedChartData', () => {
	it('uses only the values cached in the chart part', () => {
		const data = cachedChartData(parseChartXml(loadFixture('line-stacked-markers.xml')));
		expect(data.plots[0].series[0]).toMatchObject({
			name: 'A',
			categories: ['2011', '2012', '2013', '2014', '2015'],
			categoriesNumeric: true,
			values: [2.4, 4.4, 1.8, 2.8, 5.6],
			live: false,
		});
		const multi = cachedChartData({
			plots: [
				{
					type: 'bar',
					axisIds: [],
					series: [
						{
							idx: 0,
							order: 0,
							categories: {
								kind: 'multiLvlStr',
								cache: ['a', 'b'],
								levels: [
									['X', null],
									['a', 'b'],
								],
							},
							values: { kind: 'num', cache: [1, null] },
						},
					],
				},
			],
			axes: [],
			displayBlanksAs: 'zero',
		});
		expect(multi.plots[0].series[0].categoryLevels).toEqual([['X', '']]);
		expect(multi.plots[0].series[0].values).toEqual([1, 0]);
	});
});

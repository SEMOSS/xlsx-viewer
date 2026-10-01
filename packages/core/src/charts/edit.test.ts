import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart';
import { loadFixture, makeWorkbook } from './__tests__/helpers';
import { createChartFromRange } from './data/create-chart';
import {
	applyQuickLayout,
	chartElementState,
	chartTypeTargetOf,
	convertChartType,
	setChartElement,
	switchRowColumn,
} from './edit';
import { parseChartXml } from './parse/parse-chart';
import { writeChartXml } from './write/write-chart';

const wb = makeWorkbook({
	Data: [
		[null, 'Revenue', 'Cost'],
		['Q1', 10, 4],
		['Q2', 20, 8],
		['Q3', 30, 12],
		['Q4', 40, 16],
	],
	Other: [[1]],
});
const all = { top: 0, left: 0, bottom: 4, right: 2 };
const column = (): ChartModel => createChartFromRange(wb, 0, all, 'column', { title: 'Sales' });
const frozen = (model: ChartModel) => {
	const json = JSON.stringify(model);
	return () => expect(JSON.stringify(model)).toBe(json);
};

describe('setChartElement', () => {
	it('adds, moves and removes the chart title', () => {
		const source = column();
		const unchanged = frozen(source);
		const none = setChartElement(source, 'chartTitle', 'none');
		expect(none.title).toBeUndefined();
		expect(none.autoTitleDeleted).toBe(true);
		expect(none.dirty).toBe(true);
		const above = setChartElement(none, 'chartTitle', 'above');
		expect(above.title).toMatchObject({ overlay: false, style: { size: 14 } });
		expect(above.autoTitleDeleted).toBe(false);
		expect(setChartElement(source, 'chartTitle', 'overlay').title).toMatchObject({
			text: 'Sales',
			overlay: true,
		});
		unchanged();
	});

	it('positions and removes the legend', () => {
		const source = column();
		source.legend = { position: 'b', layout: { x: 0.1, y: 0.1 } };
		const right = setChartElement(source, 'legend', 'right');
		expect(right.legend).toEqual({ position: 'r', overlay: false });
		expect(setChartElement(source, 'legend', 'top').legend?.position).toBe('t');
		expect(setChartElement(source, 'legend', 'left').legend?.position).toBe('l');
		expect(setChartElement(source, 'legend', 'none').legend).toBeUndefined();
		const fresh = setChartElement(setChartElement(source, 'legend', 'none'), 'legend', 'bottom');
		expect(fresh.legend).toMatchObject({ position: 'b', textStyle: { size: 9 } });
	});

	it('shows values, or percentages on pies, as data labels', () => {
		const bars = setChartElement(column(), 'dataLabels', 'show');
		expect(
			bars.plots[0].series.every((s) => s.dataLabels?.showValue && !s.dataLabels.showPercent),
		).toBe(true);
		const pie = setChartElement(convertChartType(column(), 'pie'), 'dataLabels', 'show');
		expect(pie.plots[0].series[0].dataLabels).toMatchObject({
			showValue: false,
			showPercent: true,
		});
		expect(
			setChartElement(bars, 'dataLabels', 'none').plots[0].series[0].dataLabels,
		).toBeUndefined();
	});

	it('toggles the primary major gridlines', () => {
		const off = setChartElement(column(), 'gridlines', 'off');
		expect(off.axes[1].majorGridlines).toBeUndefined();
		expect(setChartElement(off, 'gridlines', 'on').axes[1].majorGridlines).toMatchObject({
			width: 9525,
		});
		const scatter = setChartElement(parseChartXml(loadFixture('scatter.xml')), 'gridlines', 'on');
		expect(scatter.axes.find((a) => a.position === 'l')?.majorGridlines).toBeDefined();
		expect(scatter.axes.find((a) => a.position === 'b')?.majorGridlines).toBeUndefined();
	});

	it('adds rotated axis titles and leaves pies alone', () => {
		const titled = setChartElement(column(), 'axisTitles', 'on');
		expect(titled.axes[0].title).toEqual({
			overlay: false,
			style: { size: 10, bold: false, color: expect.any(Object) },
		});
		expect(titled.axes[1].title?.rotation).toBe(-90);
		expect(setChartElement(titled, 'axisTitles', 'off').axes.every((a) => !a.title)).toBe(true);
		const pie = convertChartType(column(), 'pie');
		expect(setChartElement(pie, 'axisTitles', 'on')).toBe(pie);
		expect(setChartElement(pie, 'gridlines', 'on')).toBe(pie);
	});

	it('writes valid chart parts', () => {
		let chart = column();
		chart = setChartElement(chart, 'axisTitles', 'on');
		chart = setChartElement(chart, 'dataLabels', 'show');
		const back = parseChartXml(writeChartXml(chart));
		expect(back.axes[1].title?.rotation).toBe(-90);
		expect(back.plots[0].series[0].dataLabels?.showValue).toBe(true);
	});
});

describe('chartElementState', () => {
	it('reports the current element options', () => {
		expect(chartElementState(column())).toEqual({
			chartTitle: 'above',
			legend: 'bottom',
			dataLabels: 'none',
			gridlines: 'on',
			axisTitles: 'off',
		});
		const changed = applyQuickLayout(column(), 2);
		expect(chartElementState(changed)).toEqual({
			chartTitle: 'above',
			legend: 'top',
			dataLabels: 'show',
			gridlines: 'off',
			axisTitles: 'off',
		});
		expect(chartElementState(convertChartType(column(), 'pie'))).toMatchObject({
			gridlines: 'off',
			axisTitles: 'off',
		});
	});
});

describe('applyQuickLayout', () => {
	it('applies a preset of elements', () => {
		const l2 = applyQuickLayout(column(), 2);
		expect(l2.title).toBeDefined();
		expect(l2.legend?.position).toBe('t');
		expect(l2.plots[0].series[0].dataLabels?.showValue).toBe(true);
		expect(l2.axes[1].majorGridlines).toBeUndefined();
		const l5 = applyQuickLayout(column(), 5);
		expect(l5.legend).toBeUndefined();
		expect(l5.axes.every((a) => a.title)).toBe(true);
		const pie = applyQuickLayout(convertChartType(column(), 'pie'), 4);
		expect(pie.title).toBeUndefined();
		expect(pie.dirty).toBe(true);
	});
});

describe('switchRowColumn', () => {
	it('plots rows as series and back again', () => {
		const source = column();
		const unchanged = frozen(source);
		const rows = switchRowColumn(source, wb, 0);
		expect(rows).not.toBeNull();
		const series = rows!.plots[0].series;
		expect(series.map((s) => s.name?.text)).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
		expect(series[0].values?.ref).toBe('Data!$B$2:$C$2');
		expect(series[0].categories?.cache).toEqual(['Revenue', 'Cost']);
		expect(rows!.title?.text).toBe('Sales');
		expect(rows!.plots[0].gapWidth).toBe(219);
		expect(rows!.dirty).toBe(true);
		const back = switchRowColumn(rows!, wb, 0)!;
		expect(back.plots[0].series.map((s) => s.values?.ref)).toEqual(
			source.plots[0].series.map((s) => s.values?.ref),
		);
		unchanged();
	});

	it('keeps the chart type formatting of the new series', () => {
		const markers = convertChartType(column(), 'lineMarkers');
		const switched = switchRowColumn(markers, wb, 0)!;
		expect(chartTypeTargetOf(switched)).toBe('lineMarkers');
		expect(switched.plots[0].series[3].marker?.fill).toEqual({
			type: 'solid',
			color: { kind: 'scheme', value: 'accent4' },
		});
	});

	it('returns null when the references cannot be switched', () => {
		expect(switchRowColumn(parseChartXml(loadFixture('combo-bar-line.xml')), wb, 0)).toBeNull();
		const scattered = column();
		scattered.plots[0].series[1].values = { kind: 'num', ref: 'Data!$C$2:$C$5' };
		scattered.plots[0].series[1].name = { ref: 'Data!$E$1' };
		expect(switchRowColumn(scattered, wb, 0)).toBeNull();
		const twoSheets = column();
		twoSheets.plots[0].series[1].values = { kind: 'num', ref: 'Other!$A$1:$A$4' };
		expect(switchRowColumn(twoSheets, wb, 0)).toBeNull();
		const unreadable = column();
		unreadable.plots[0].series[0].values = { kind: 'num', ref: '[1]Data!$B$2:$B$5' };
		expect(switchRowColumn(unreadable, wb, 0)).toBeNull();
		expect(switchRowColumn(parseChartXml(loadFixture('stock.xml')), wb, 0)).toBeNull();
	});
});

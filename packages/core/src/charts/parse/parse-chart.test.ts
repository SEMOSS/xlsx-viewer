import { describe, expect, it } from 'vitest';
import type { XlsxWarning } from '../../types/warnings';
import { fixtureNames, loadFixture } from '../__tests__/helpers';
import { parseChartXml } from './parse-chart';

const parse = (name: string) => parseChartXml(loadFixture(name));

describe('parseChartXml: every fixture', () => {
	it.each(fixtureNames())('%s keeps the raw XML and is clean', (name) => {
		const xml = loadFixture(name);
		const chart = parseChartXml(xml);
		expect(chart.rawXml).toBe(xml);
		expect(chart.dirty).toBe(false);
		expect(chart.plots.length).toBeGreaterThan(0);
		for (const plot of chart.plots) {
			expect(plot.series.length).toBeGreaterThan(0);
			for (const s of plot.series) {
				expect(s.values?.cache?.length ?? 0).toBeGreaterThan(0);
			}
		}
	});

	it('reads plot types, series counts and axes of each fixture', () => {
		const summary = Object.fromEntries(
			fixtureNames().map((name) => {
				const c = parse(name);
				return [
					name,
					`${c.plots.map((p) => `${p.type}:${p.series.length}`).join('+')} ${c.axes.map((a) => a.kind).join(',')}`,
				];
			}),
		);
		expect(summary).toEqual({
			'area.xml': 'area:2 cat,val',
			'bar-filtered-series.xml': 'bar:2 cat,val',
			'bar-manual-layout.xml': 'bar:1 cat,val',
			'bar3d-horizontal.xml': 'bar:2 cat,val,ser',
			'bubble-office.xml': 'bubble:3 val,val',
			'combo-bar-line.xml': 'bar:1+line:2 cat,val',
			'doughnut-office.xml': 'doughnut:1 ',
			'line-callouts-office.xml': 'line:3 cat,val',
			'line-stacked-markers.xml': 'line:2 cat,val',
			'pie-office.xml': 'pie:1 ',
			'pie-title-runs.xml': 'pie:1 ',
			'radar-office.xml': 'radar:2 cat,val',
			'scatter.xml': 'scatter:2 val,val',
			'stock.xml': 'stock:4 cat,val',
			'surface.xml': 'surface:3 cat,val,ser',
		});
	});
});

describe('parseChartXml: series data', () => {
	it('reads references, caches and format codes', () => {
		const [s0, s1] = parse('area.xml').plots[0].series;
		expect(s0.name).toEqual({ ref: 'Sheet1!$B$1', text: 'Revenue' });
		expect(s0.categories).toEqual({
			kind: 'str',
			ref: 'Sheet1!$A$2:$A$5',
			cache: ['Q1', 'Q2', 'Q3', 'Q4'],
		});
		expect(s0.values).toEqual({
			kind: 'num',
			ref: 'Sheet1!$B$2:$B$5',
			cache: [45, 62, 58, 71],
			formatCode: 'General',
		});
		expect(s1.values?.cache).toEqual([30, 41, 38, 52]);
	});

	it('reads numeric categories, X values and bubble sizes', () => {
		const lines = parse('line-stacked-markers.xml').plots[0];
		expect(lines.series[0].categories).toMatchObject({
			kind: 'num',
			cache: [2011, 2012, 2013, 2014, 2015],
		});
		const bubble = parse('bubble-office.xml').plots[0];
		expect(bubble.series[1].categories?.cache).toEqual([1, 2, 3]);
		expect(bubble.series[1].values?.cache).toEqual([3, 6, 9]);
		expect(bubble.series[1].bubbleSizes?.cache).toEqual([7, 3, 12]);
		expect(bubble.bubbleScale).toBe(100);
	});

	it('keeps series idx / order and skips filtered (hidden) series', () => {
		const plot = parse('bar-filtered-series.xml').plots[0];
		expect(plot.series.map((s) => [s.idx, s.order, s.name?.text])).toEqual([
			[0, 0, 'Series A'],
			[2, 2, 'Series C'],
		]);
	});

	it('reads gap-filled caches by point index', () => {
		const xml = loadFixture('area.xml').replace(
			'<c:ptCount val="4"/><c:pt idx="0"><c:v>45</c:v></c:pt><c:pt idx="1"><c:v>62</c:v></c:pt>',
			'<c:ptCount val="4"/><c:pt idx="1"><c:v>62</c:v></c:pt>',
		);
		expect(parseChartXml(xml).plots[0].series[0].values?.cache).toEqual([null, 62, 58, 71]);
	});

	it('reads multi-level categories (outer level first) and literals', () => {
		const xml = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
			<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Lit</c:v></c:tx>
			<c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/>
				<c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Jan</c:v></c:pt><c:pt idx="3"><c:v>Feb</c:v></c:pt></c:lvl>
				<c:lvl><c:pt idx="0"><c:v>2023</c:v></c:pt><c:pt idx="2"><c:v>2024</c:v></c:pt></c:lvl>
			</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
			<c:val><c:numLit><c:formatCode>0.0</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:numLit></c:val>
			</c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
		const s = parseChartXml(xml).plots[0].series[0];
		expect(s.name).toEqual({ text: 'Lit' });
		expect(s.categories?.kind).toBe('multiLvlStr');
		expect(s.categories?.levels).toEqual([
			['2023', null, '2024', null],
			['Jan', 'Feb', 'Jan', 'Feb'],
		]);
		expect(s.categories?.cache).toEqual(['Jan', 'Feb', 'Jan', 'Feb']);
		expect(s.values).toEqual({ kind: 'literal', cache: [1.5, null, null, 4], formatCode: '0.0' });
	});
});

describe('parseChartXml: formatting', () => {
	it('reads scheme colors with transforms and outlines', () => {
		const s = parse('bubble-office.xml').plots[0].series[0];
		expect(s.fill).toEqual({
			type: 'solid',
			color: { kind: 'scheme', value: 'accent1', transforms: [{ name: 'alpha', val: 75000 }] },
		});
		expect(s.line).toEqual({ fill: { type: 'none' } });
		const point = parse('bar-manual-layout.xml').plots[0].series[0].points?.[2];
		expect(point?.fill).toEqual({
			type: 'solid',
			color: {
				kind: 'scheme',
				value: 'accent1',
				transforms: [
					{ name: 'lumMod', val: 60000 },
					{ name: 'lumOff', val: 40000 },
				],
			},
		});
	});

	it('reads markers, smoothing and line caps', () => {
		const s = parse('line-callouts-office.xml').plots[0].series[0];
		expect(s.line).toMatchObject({ width: 28575, cap: 'rnd' });
		expect(s.marker).toMatchObject({ symbol: 'circle', size: 5 });
		expect(s.smooth).toBe(false);
		const auto = parse('line-stacked-markers.xml').plots[0].series[0].marker;
		expect(auto?.symbol).toBeUndefined();
		expect(auto?.fill).toEqual({ type: 'solid', color: { kind: 'srgb', value: '00B0F0' } });
	});

	it('reads data labels with per-point overrides and label boxes', () => {
		const labels = parse('line-callouts-office.xml').plots[0].series[0].dataLabels;
		expect(labels).toMatchObject({
			showValue: true,
			showCategoryName: false,
			showLeaderLines: false,
		});
		expect(labels?.fill).toEqual({ type: 'solid', color: { kind: 'srgb', value: 'FFFFCC' } });
		expect(labels?.points?.[0]).toMatchObject({
			idx: 1,
			showValue: true,
			layout: { x: 0.056818181818181816, y: -0.12 },
		});
		const pie = parse('pie-office.xml').plots[0].series[0].dataLabels;
		expect(pie).toMatchObject({ showCategoryName: true, showPercent: true, separator: ', ' });
	});

	it('reads chart, plot and title formatting', () => {
		const chart = parse('pie-office.xml');
		expect(chart.style).toBe(102);
		expect(chart.roundedCorners).toBe(false);
		expect(chart.chartFill).toEqual({ type: 'none' });
		expect(chart.title).toMatchObject({ overlay: false });
		expect(chart.title?.text).toBeUndefined();
		expect(chart.plots[0]).toMatchObject({ varyColors: true, firstSliceAngle: 0 });
		expect(chart.legend).toMatchObject({ position: 'b', overlay: false });
		expect(parse('doughnut-office.xml').plots[0].holeSize).toBe(75);
	});

	it('defaults rounded corners to true when the element is absent (schema default)', () => {
		expect(parse('area.xml').roundedCorners).toBe(true);
	});

	it('reads rich title runs', () => {
		const title = parse('pie-title-runs.xml').title;
		expect(title?.text).toBe('Sales Overview');
		expect(title?.paragraphs?.[0].runs).toEqual([
			{ text: 'Sales ', style: { bold: true } },
			{ text: 'Overview', style: { italic: true, color: { kind: 'srgb', value: 'FF0000' } } },
		]);
	});

	it('reads the manual plot area layout and axis formatting', () => {
		const chart = parse('bar-manual-layout.xml');
		expect(chart.plotAreaLayout).toMatchObject({
			layoutTarget: 'inner',
			xMode: 'edge',
			yMode: 'edge',
		});
		expect(chart.plotAreaLayout?.h).toBeCloseTo(0.8606, 3);
		const [cat, val] = chart.axes;
		expect(cat).toMatchObject({
			id: -849514352,
			kind: 'cat',
			position: 'b',
			crossAxisId: -849644336,
			majorTickMark: 'out',
		});
		expect(val).toMatchObject({
			kind: 'val',
			numFmt: { code: '0%', sourceLinked: true },
			crossBetween: 'between',
		});
		expect(val.majorGridlines).toMatchObject({ width: 9525 });
		expect(val.line).toEqual({ fill: { type: 'none' } });
		expect(val.textStyle?.size).toBeCloseTo(11.95, 2);
	});

	it('reads plot-specific settings', () => {
		const combo = parse('combo-bar-line.xml');
		expect(combo.plots[0]).toMatchObject({
			barDirection: 'col',
			grouping: 'clustered',
			gapWidth: 150,
			overlap: -27,
		});
		expect(combo.plots[1]).toMatchObject({ grouping: 'standard', marker: true });
		expect(combo.plots[0].axisIds).toEqual(combo.plots[1].axisIds);
		expect(parse('bar3d-horizontal.xml').plots[0]).toMatchObject({
			is3D: true,
			barDirection: 'bar',
		});
		expect(parse('scatter.xml').plots[0].scatterStyle).toBe('marker');
		expect(parse('radar-office.xml').plots[0].radarStyle).toBe('marker');
		expect(parse('stock.xml').plots[0]).toMatchObject({ hiLowLines: true, upDownBars: true });
		expect(parse('line-stacked-markers.xml')).toMatchObject({
			displayBlanksAs: 'zero',
			plotVisibleOnly: true,
		});
		expect(parse('line-stacked-markers.xml').legend?.textStyle).toEqual({ size: 18 });
	});
});

describe('parseChartXml: robustness', () => {
	it('keeps chartex parts as raw XML and warns', () => {
		const xml =
			'<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData/><cx:chart><cx:title><cx:tx><cx:rich><a:p><a:r><a:t>Funnel</a:t></a:r></a:p></cx:rich></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>';
		const warnings: XlsxWarning[] = [];
		const chart = parseChartXml(xml, {
			partPath: 'xl/charts/chartEx1.xml',
			onWarning: (w) => warnings.push(w),
		});
		expect(chart).toMatchObject({
			plots: [],
			axes: [],
			rawXml: xml,
			dirty: false,
			title: { text: 'Funnel' },
		});
		expect(warnings).toEqual([
			expect.objectContaining({ code: 'chart-unsupported', part: 'xl/charts/chartEx1.xml' }),
		]);
	});

	it('survives malformed XML', () => {
		const warnings: XlsxWarning[] = [];
		const chart = parseChartXml('<c:chartSpace><c:chart>', { onWarning: (w) => warnings.push(w) });
		expect(chart).toEqual({ plots: [], axes: [], rawXml: '<c:chartSpace><c:chart>', dirty: false });
		expect(warnings[0].code).toBe('chart-unreadable');
	});

	it('warns about approximated chart types', () => {
		const warnings: XlsxWarning[] = [];
		parseChartXml(loadFixture('surface.xml'), { onWarning: (w) => warnings.push(w) });
		parseChartXml(loadFixture('bar3d-horizontal.xml'), { onWarning: (w) => warnings.push(w) });
		expect(warnings.map((w) => w.code)).toEqual([
			'chart-type-approximated',
			'chart-type-approximated',
		]);
	});

	it('accepts unprefixed (default namespace) chart XML', () => {
		const xml = loadFixture('area.xml')
			.replace(/<c:/g, '<')
			.replace(/<\/c:/g, '</')
			.replace('xmlns:c=', 'xmlns=');
		const chart = parseChartXml(xml);
		expect(chart.plots[0].type).toBe('area');
		expect(chart.plots[0].series[1].values?.cache).toEqual([30, 41, 38, 52]);
	});
});

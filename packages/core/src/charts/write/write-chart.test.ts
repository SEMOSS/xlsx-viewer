import { describe, expect, it } from 'vitest';
import type { ChartDataLabels, ChartModel } from '../../types/chart';
import { isElement, localName } from '../../xml/dom';
import { parseXmlRoot } from '../../xml/parser';
import type { XmlElement } from '../../xml/types';
import { fixtureNames, loadFixture } from '../__tests__/helpers';
import { parseChartXml } from '../parse/parse-chart';
import { CHART_ORDER, seriesOrderKey } from './schema';
import { writeChartXml } from './write-chart';

const SCALING_ORDER = ['logBase', 'orientation', 'max', 'min', 'extLst'];

/** Collect schema-order violations of a generated chart part. */
function orderErrors(root: XmlElement): string[] {
	const errors: string[] = [];
	const visit = (el: XmlElement, container: string, path: string): void => {
		const local = localName(el.name);
		if (local === 'extLst' || local === 'AlternateContent') {
			return;
		}
		const key = local === 'ser' ? seriesOrderKey(container) : local;
		const order = local === 'scaling' ? SCALING_ORDER : CHART_ORDER[key];
		const kids = el.children.filter(isElement);
		// Plot area: the chart groups and the axes are each an unordered choice.
		const rankOf = (name: string): number =>
			local === 'plotArea' && /Chart$/.test(name)
				? order!.indexOf('areaChart')
				: local === 'plotArea' && /Ax$/.test(name)
					? order!.indexOf('valAx')
					: order!.indexOf(name);
		if (order) {
			let last = -1;
			for (const k of kids) {
				const rank = rankOf(localName(k.name));
				if (rank < 0) {
					errors.push(`${path}/${k.name}: not allowed in ${key}`);
				} else if (rank < last) {
					errors.push(`${path}/${k.name}: out of order in ${key}`);
				}
				last = Math.max(last, rank);
			}
		}
		for (const k of kids) {
			visit(k, /Chart$/.test(local) ? local : container, `${path}/${localName(k.name)}`);
		}
	};
	visit(root, '', '');
	return errors;
}

/** Fill in values the writer always emits, so models can be compared after a round trip. */
function normalize(model: ChartModel): unknown {
	const copy = JSON.parse(JSON.stringify(model)) as ChartModel;
	delete copy.rawXml;
	delete copy.dirty;
	const labels = (l: ChartDataLabels | undefined) => {
		if (!l || l.deleted) {
			return;
		}
		for (const key of [
			'showLegendKey',
			'showValue',
			'showCategoryName',
			'showSeriesName',
			'showPercent',
			'showBubbleSize',
		] as const) {
			l[key] = Boolean(l[key]);
		}
		for (const p of l.points ?? []) {
			for (const key of [
				'showLegendKey',
				'showValue',
				'showCategoryName',
				'showSeriesName',
				'showPercent',
				'showBubbleSize',
			] as const) {
				p[key] = Boolean(p[key] ?? l[key]);
			}
		}
	};
	for (const plot of copy.plots) {
		if (plot.type === 'bar') {
			plot.grouping ??= 'clustered';
		}
		if (plot.type === 'line' || plot.type === 'area') {
			plot.grouping ??= 'standard';
		}
		if (plot.type === 'scatter') {
			plot.scatterStyle ??= 'lineMarker';
		}
		if (plot.type === 'radar') {
			plot.radarStyle ??= 'marker';
		}
		labels(plot.dataLabels);
		plot.series.forEach((s) => labels(s.dataLabels));
	}
	for (const axis of copy.axes) {
		if (axis.numFmt) {
			axis.numFmt.sourceLinked = Boolean(axis.numFmt.sourceLinked);
		}
	}
	return copy;
}

const roundTrip = (model: ChartModel) => parseChartXml(writeChartXml({ ...model, dirty: true }));

describe('writeChartXml', () => {
	it('returns the original XML for unedited charts', () => {
		const xml = loadFixture('pie-office.xml');
		expect(writeChartXml(parseChartXml(xml))).toBe(xml);
	});

	it.each(fixtureNames())('round-trips %s through the model', (name) => {
		const original = parseChartXml(loadFixture(name));
		const xml = writeChartXml({ ...original, dirty: true });
		expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
		expect(orderErrors(parseXmlRoot(xml))).toEqual([]);
		expect(normalize(parseChartXml(xml))).toEqual(normalize(original));
	});

	it('is stable: writing a written chart again gives the same XML', () => {
		for (const name of fixtureNames()) {
			const once = writeChartXml({ ...parseChartXml(loadFixture(name)), dirty: true });
			const twice = writeChartXml({ ...parseChartXml(once), dirty: true });
			expect(twice).toBe(once);
		}
	});

	it('keeps markup the model does not cover when regenerating', () => {
		const xml = writeChartXml({
			...parseChartXml(loadFixture('bar-filtered-series.xml')),
			dirty: true,
		});
		expect(xml).toContain('<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>');
		expect(xml).toContain('<c15:filteredBarSeries>');
		expect(xml).toContain('c16:uniqueId');
		expect(xml).toContain('<c:lblOffset val="100"/>');
		expect(xml).toContain(
			'xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"',
		);
		expect(xml).toContain('<c:showDLblsOverMax val="0"/>');
		const shapes = writeChartXml({ ...parseChartXml(loadFixture('pie-office.xml')), dirty: true });
		expect(shapes).toContain('<c:leaderLines>');
		const user = loadFixture('area.xml').replace(
			'</c:chartSpace>',
			'<c:userShapes r:id="rId9"/></c:chartSpace>',
		);
		expect(writeChartXml({ ...parseChartXml(user), dirty: true })).toContain(
			'<c:userShapes r:id="rId9"/>',
		);
	});

	it('keeps the 3-D view of 3-D charts', () => {
		const xml = writeChartXml({
			...parseChartXml(loadFixture('bar3d-horizontal.xml')),
			dirty: true,
		});
		expect(xml).toContain('<c:bar3DChart>');
		expect(xml).toContain('<c:shape val="box"/>');
		expect(xml).toContain('<c:view3D>');
		expect(xml.match(/<c:axId /g)?.length).toBe(3 + 3);
	});

	it('writes the c14 style through mc:AlternateContent', () => {
		const xml = writeChartXml({ ...parseChartXml(loadFixture('pie-office.xml')), dirty: true });
		expect(xml).toContain('<c14:style val="102"/>');
		expect(xml).toContain('<mc:Fallback><c:style val="2"/></mc:Fallback>');
	});

	it('applies model edits', () => {
		const chart = parseChartXml(loadFixture('combo-bar-line.xml'));
		chart.plots[0].series[0].fill = { type: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
		chart.title = { text: 'Edited & <escaped>' };
		chart.legend = { position: 'r' };
		chart.axes[1].min = 10;
		chart.axes[1].majorUnit = 5;
		chart.dirty = true;
		const xml = writeChartXml(chart);
		expect(xml).toContain('Edited &amp; &lt;escaped&gt;');
		const back = parseChartXml(xml);
		expect(back.plots[0].series[0].fill).toEqual({
			type: 'solid',
			color: { kind: 'srgb', value: 'FF0000' },
		});
		expect(back.title?.text).toBe('Edited & <escaped>');
		expect(back.legend?.position).toBe('r');
		expect(back.axes[1]).toMatchObject({ min: 10, majorUnit: 5 });
	});

	it('creates default axes for cartesian plots without them', () => {
		const chart: ChartModel = {
			plots: [
				{
					type: 'bar',
					barDirection: 'bar',
					series: [
						{
							idx: 0,
							order: 0,
							values: { kind: 'literal', cache: [1, 2, 3], formatCode: 'General' },
						},
					],
					axisIds: [],
				},
			],
			axes: [],
			dirty: true,
		};
		const back = roundTrip(chart);
		expect(back.axes.map((a) => `${a.kind}@${a.position}`)).toEqual(['cat@l', 'val@b']);
		expect(back.plots[0].axisIds).toEqual(back.axes.map((a) => a.id));
		expect(back.plots[0].series[0].values).toEqual({
			kind: 'literal',
			cache: [1, 2, 3],
			formatCode: 'General',
		});
	});

	it('falls back to valid containers when a model cannot be written as-is', () => {
		const stock = parseChartXml(loadFixture('stock.xml'));
		stock.plots[0].series = stock.plots[0].series.slice(0, 2);
		expect(writeChartXml({ ...stock, dirty: true })).toContain('<c:lineChart>');
		const line3d: ChartModel = {
			plots: [{ type: 'line', is3D: true, series: [{ idx: 0, order: 0 }], axisIds: [] }],
			axes: [],
			dirty: true,
		};
		expect(writeChartXml(line3d)).toContain('<c:lineChart>');
	});

	it('never rewrites chartex parts', () => {
		const xml =
			'<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData/></cx:chartSpace>';
		const chart = parseChartXml(xml);
		expect(writeChartXml({ ...chart, dirty: true })).toBe(xml);
	});

	it('writes literal, string and multi-level data sources', () => {
		const chart: ChartModel = {
			plots: [
				{
					type: 'bar',
					barDirection: 'col',
					series: [
						{
							idx: 0,
							order: 0,
							name: { text: 'Literal name' },
							categories: {
								kind: 'multiLvlStr',
								ref: 'Sheet1!$A$2:$B$3',
								cache: ['Jan', 'Feb'],
								levels: [
									['2024', null],
									['Jan', 'Feb'],
								],
							},
							values: {
								kind: 'num',
								ref: "'My Sheet'!$C$2:$C$3",
								cache: [1e-7, 2.5],
								formatCode: '0.00',
							},
						},
						{
							idx: 1,
							order: 1,
							categories: { kind: 'literal', cache: ['a', 'b'] },
							values: { kind: 'literal', cache: [3, null], formatCode: 'General' },
						},
					],
					axisIds: [],
				},
			],
			axes: [],
			dirty: true,
		};
		const xml = writeChartXml(chart);
		expect(xml).toContain('<c:v>1E-7</c:v>');
		expect(xml).toContain('<c:f>&apos;My Sheet&apos;!$C$2:$C$3</c:f>'.replace(/&apos;/g, "'"));
		const back = roundTrip(chart);
		expect(back.plots[0].series[0].categories).toEqual(chart.plots[0].series[0].categories);
		expect(back.plots[0].series[0].values).toEqual(chart.plots[0].series[0].values);
		expect(back.plots[0].series[1].categories).toEqual({ kind: 'literal', cache: ['a', 'b'] });
		expect(back.plots[0].series[1].values).toEqual({
			kind: 'literal',
			cache: [3, null],
			formatCode: 'General',
		});
		expect(back.plots[0].series[0].name).toEqual({ text: 'Literal name' });
	});
});

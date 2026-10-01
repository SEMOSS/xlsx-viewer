import { describe, expect, it, vi } from 'vitest';
import { createDefaultTheme } from '../../styles/defaults';
import type { ChartModel } from '../../types/chart';
import {
	assertFinite,
	byRole,
	fixtureNames,
	literalChart,
	loadFixture,
	texts,
} from '../__tests__/helpers';
import { cachedChartData } from '../data/resolve-data';
import { parseChartXml } from '../parse/parse-chart';
import { chartTitleText, describeChart } from './describe';
import { layoutChart } from './layout-chart';
import type { ChartRectNode, ChartRenderTree, ChartTextNode } from './types';

const theme = createDefaultTheme();
const render = (chart: ChartModel, w = 480, h = 288): ChartRenderTree => {
	const tree = layoutChart(chart, cachedChartData(chart), theme, w, h);
	assertFinite(tree);
	return tree;
};
const two = (model: Partial<ChartModel> = {}) =>
	literalChart(
		{ type: 'bar', barDirection: 'col' },
		[
			[1, 2],
			[3, 4],
		],
		{ categories: ['a', 'b'], names: ['Alpha', 'Beta'], model },
	);
const textNodes = (tree: ChartRenderTree, role: string) =>
	byRole(tree, role).filter((n): n is ChartTextNode => n.type === 'text');
const plotBox = (tree: ChartRenderTree) => {
	const b = byRole(tree, 'bar') as ChartRectNode[];
	return {
		left: Math.min(...b.map((r) => r.x)),
		right: Math.max(...b.map((r) => r.x + r.width)),
		top: Math.min(...b.map((r) => r.y)),
		bottom: Math.max(...b.map((r) => r.y + r.height)),
	};
};

describe('layoutChart: fixtures', () => {
	it.each(fixtureNames())('lays out %s', (name) => {
		const chart = parseChartXml(loadFixture(name));
		for (const [w, h] of [
			[480, 288],
			[160, 100],
			[1200, 700],
			[1, 1],
			[0, 0],
		]) {
			const tree = render(chart, w, h);
			expect(tree.width).toBe(Math.max(1, w));
			expect(tree.root.children[0]).toMatchObject({ type: 'rect', role: 'chart-area' });
			expect(tree.description.length).toBeGreaterThan(0);
		}
	});
});

describe('layoutChart: chart area and title', () => {
	it('fills the chart area white with a light border by default', () => {
		const area = render(two()).root.children[0] as ChartRectNode;
		expect(area).toMatchObject({ fill: '#FFFFFF', stroke: '#D9D9D9', strokeWidth: 1 });
		expect(area.rx).toBeUndefined();
		const rounded = render(two({ roundedCorners: true })).root.children[0] as ChartRectNode;
		expect(rounded.rx).toBeGreaterThan(0);
		const bare = render(two({ chartFill: { type: 'none' }, chartLine: { fill: { type: 'none' } } }))
			.root.children[0] as ChartRectNode;
		expect(bare.fill).toBeUndefined();
		expect(bare.stroke).toBeUndefined();
	});

	it('shows automatic titles like Excel', () => {
		expect(texts(render(two({ title: {} })), 'title')).toEqual(['Chart Title']);
		const single = literalChart({ type: 'bar', barDirection: 'col' }, [[1]], {
			names: ['Revenue'],
			model: { title: {} },
		});
		expect(texts(render(single), 'title')).toEqual(['Revenue']);
		expect(texts(render(two()), 'title')).toEqual([]);
		const title = textNodes(render(two({ title: { text: 'Sales' } })), 'title')[0];
		expect(title).toMatchObject({ anchor: 'middle', x: 240 });
		expect(title.fontSize).toBeCloseTo((14 * 96) / 72);
	});

	it('keeps rich title runs and moves the plot below the title', () => {
		const tree = render(parseChartXml(loadFixture('pie-title-runs.xml')));
		const title = textNodes(tree, 'title')[0];
		expect(title.runs).toEqual([
			[
				{ text: 'Sales ', fontWeight: 'bold' },
				{ text: 'Overview', fontStyle: 'italic', fill: '#FF0000' },
			],
		]);
		const withTitle = plotBox(render(two({ title: { text: 'T' } })));
		const without = plotBox(render(two()));
		expect(withTitle.top).toBeGreaterThan(without.top);
		const overlay = plotBox(render(two({ title: { text: 'T', overlay: true } })));
		expect(overlay.top).toBeCloseTo(without.top, 5);
	});

	it('prefers live titles from linked cells', () => {
		const chart = two({ title: { ref: 'Sheet1!$A$1', text: 'Cached' } });
		const data = { ...cachedChartData(chart), titles: { chart: 'Live' } };
		expect(texts(layoutChart(chart, data, theme, 480, 288), 'title')).toEqual(['Live']);
	});

	it('wraps long titles', () => {
		const tree = render(
			two({
				title: {
					text: 'A very long chart title that cannot fit on a single line of a small chart',
				},
			}),
			240,
			200,
		);
		expect(textNodes(tree, 'title')[0].lines.length).toBeGreaterThan(1);
	});
});

describe('layoutChart: legend', () => {
	it('places the legend on each side and shrinks the plot area', () => {
		const none = plotBox(render(two()));
		const right = render(two({ legend: { position: 'r' } }));
		const rightLabels = textNodes(right, 'legend-entry');
		expect(rightLabels.map((t) => t.lines[0])).toEqual(['Alpha', 'Beta']);
		expect(Math.min(...rightLabels.map((t) => t.x))).toBeGreaterThan(plotBox(right).right);
		expect(plotBox(right).right).toBeLessThan(none.right);
		expect(rightLabels[1].y).toBeGreaterThan(rightLabels[0].y);
		const bottom = render(two({ legend: { position: 'b' } }));
		const bottomLabels = textNodes(bottom, 'legend-entry');
		expect(bottomLabels[0].y).toBeCloseTo(bottomLabels[1].y, 5);
		expect(bottomLabels[0].y).toBeGreaterThan(plotBox(bottom).bottom);
		const top = render(two({ legend: { position: 't' } }));
		expect(textNodes(top, 'legend-entry')[0].y).toBeLessThan(plotBox(top).top);
		const left = render(two({ legend: { position: 'l' } }));
		expect(textNodes(left, 'legend-entry')[0].x).toBeLessThan(plotBox(left).left);
		const corner = render(two({ legend: { position: 'tr' } }));
		expect(textNodes(corner, 'legend-entry')[0].y).toBeLessThan(
			textNodes(right, 'legend-entry')[0].y,
		);
	});

	it('does not shrink the plot for overlay legends', () => {
		expect(plotBox(render(two({ legend: { position: 'r', overlay: true } }))).right).toBeCloseTo(
			plotBox(render(two())).right,
			5,
		);
	});

	it('drops deleted entries and reverses stacked series in vertical legends', () => {
		expect(
			texts(render(two({ legend: { position: 'r', deletedEntries: [0] } })), 'legend-entry'),
		).toEqual(['Beta']);
		const stacked = two({ legend: { position: 'r' } });
		stacked.plots[0].grouping = 'stacked';
		expect(texts(render(stacked), 'legend-entry')).toEqual(['Beta', 'Alpha']);
		stacked.legend = { position: 'b' };
		expect(texts(render(stacked), 'legend-entry')).toEqual(['Alpha', 'Beta']);
	});

	it('wraps a crowded horizontal legend into rows', () => {
		const names = Array.from({ length: 12 }, (_, i) => `Series number ${i + 1}`);
		const chart = literalChart(
			{ type: 'line', marker: false },
			names.map(() => [1, 2]),
			{ names, model: { legend: { position: 'b' } } },
		);
		const ys = new Set(textNodes(render(chart), 'legend-entry').map((t) => Math.round(t.y)));
		expect(ys.size).toBeGreaterThan(1);
	});

	it('draws line keys with markers and box keys for filled series', () => {
		const tree = render(parseChartXml(loadFixture('combo-bar-line.xml')));
		const keys = byRole(tree, 'legend-key');
		expect(keys.filter((k) => k.type === 'line')).toHaveLength(2);
		expect(keys.filter((k) => k.type === 'rect').length).toBeGreaterThanOrEqual(1);
	});
});

describe('layoutChart: plot area', () => {
	it('honours a manual inner plot area', () => {
		const chart = two({
			plotAreaLayout: {
				layoutTarget: 'inner',
				xMode: 'edge',
				yMode: 'edge',
				x: 0.25,
				y: 0.1,
				w: 0.5,
				h: 0.6,
			},
		});
		const box = plotBox(render(chart));
		expect(box.bottom).toBeCloseTo(288 * 0.7, 0);
		expect(box.left).toBeGreaterThan(480 * 0.25);
		expect(box.right).toBeLessThan(480 * 0.75);
	});

	it('shows a placeholder for charts it cannot draw', () => {
		const tree = render({ plots: [], axes: [], title: { text: 'Funnel' } });
		expect(texts(tree, 'placeholder')).toEqual(['This chart type is not supported']);
		expect(texts(tree, 'title')).toEqual(['Funnel']);
	});

	it('uses the supplied text measurer', () => {
		const measure = vi.fn((text: string, font: { size: number }) => text.length * font.size * 0.5);
		layoutChart(two({ legend: { position: 'r' } }), cachedChartData(two()), theme, 480, 288, {
			measureText: measure,
		});
		expect(measure).toHaveBeenCalled();
	});

	it('resolves theme fonts and chart-wide text sizes', () => {
		const chart = two({ textStyle: { size: 18, font: '+mn-lt' }, legend: { position: 'b' } });
		const legend = textNodes(render(chart), 'legend-entry')[0];
		expect(legend.fontSize).toBeCloseTo(24);
		expect(legend.fontFamily.startsWith("'Calibri'")).toBe(true);
	});
});

describe('describeChart / chartTitleText', () => {
	it('summarises title, type and series count', () => {
		expect(describeChart(parseChartXml(loadFixture('combo-bar-line.xml')))).toBe(
			'Combo, combo chart with 3 series',
		);
		expect(describeChart(parseChartXml(loadFixture('bar3d-horizontal.xml')))).toBe(
			'bar chart with 2 series',
		);
		const stacked = literalChart({ type: 'area', grouping: 'percentStacked' }, [[1]]);
		expect(describeChart(stacked)).toBe('100% stacked area chart with 1 series');
		expect(chartTitleText({ plots: [], axes: [] })).toBeUndefined();
	});
});

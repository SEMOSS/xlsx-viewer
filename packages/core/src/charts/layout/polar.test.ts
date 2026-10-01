import { describe, expect, it } from 'vitest';
import { createDefaultTheme } from '../../styles/defaults';
import type { ChartModel } from '../../types/chart';
import { assertFinite, byRole, literalChart, loadFixture, texts } from '../__tests__/helpers';
import { cachedChartData } from '../data/resolve-data';
import { parseChartXml } from '../parse/parse-chart';
import { layoutChart } from './layout-chart';
import { pieSlices } from './plot-pie';
import type { ChartPathNode, ChartRenderTree } from './types';

const theme = createDefaultTheme();
const render = (chart: ChartModel, w = 480, h = 288): ChartRenderTree => {
	const tree = layoutChart(chart, cachedChartData(chart), theme, w, h);
	assertFinite(tree);
	return tree;
};
const slices = (tree: ChartRenderTree) => byRole(tree, 'slice') as ChartPathNode[];

describe('pieSlices', () => {
	it('starts at the first slice angle and sweeps clockwise in proportion', () => {
		const s = pieSlices([1, 2, 1], 90);
		expect(s.map((x) => [x.start, x.end])).toEqual([
			[90, 180],
			[180, 360],
			[360, 450],
		]);
		expect(s.map((x) => x.share)).toEqual([0.25, 0.5, 0.25]);
	});

	it('skips empty points and uses absolute values', () => {
		const s = pieSlices([3, null, 0, -1], 0);
		expect(s.map((x) => x.point)).toEqual([0, 3]);
		expect(s[1].share).toBe(0.25);
		expect(pieSlices([null, 0])).toEqual([]);
	});
});

describe('pie and doughnut charts', () => {
	it('draws the first series with automatic point colors and white borders', () => {
		const chart = literalChart(
			{ type: 'pie', varyColors: true },
			[
				[1, 1, 2],
				[5, 5, 5],
			],
			{
				categories: ['a', 'b', 'c'],
				model: { legend: { position: 'r' } },
			},
		);
		const tree = render(chart);
		const s = slices(tree);
		expect(s).toHaveLength(3);
		expect(s.map((x) => x.fill)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
		expect(s.every((x) => x.stroke === '#FFFFFF')).toBe(true);
		expect(texts(tree, 'legend-entry')).toEqual(['a', 'b', 'c']);
	});

	it('uses one color without varyColors', () => {
		const chart = literalChart({ type: 'pie', varyColors: false }, [[1, 2]], {
			categories: ['a', 'b'],
		});
		expect(new Set(slices(render(chart)).map((x) => x.fill)).size).toBe(1);
	});

	it('explodes slices and shrinks the pie to fit', () => {
		const flat = slices(render(literalChart({ type: 'pie' }, [[1, 1]])));
		const exploded = literalChart({ type: 'pie' }, [[1, 1]]);
		exploded.plots[0].series[0].points = [{ idx: 0, explosion: 25 }];
		const e = slices(render(exploded));
		const radius = (d: string) => Number(/A([\d.]+)/.exec(d)?.[1]);
		expect(radius(e[0].d)).toBeCloseTo(radius(flat[0].d) / 1.25, 1);
		expect(e[0].d.split(' ')[0]).not.toBe(e[1].d.split(' ')[0]);
	});

	it('draws doughnut rings per series with the hole size', () => {
		const chart = literalChart(
			{ type: 'doughnut', holeSize: 50 },
			[
				[1, 1],
				[1, 3],
			],
			{ names: ['inner', 'outer'] },
		);
		const s = slices(render(chart));
		expect(s).toHaveLength(4);
		const radii = (d: string) => [...d.matchAll(/A([\d.]+)/g)].map((m) => Number(m[1]));
		const [outerInner, innerInner] = [radii(s[0].d), radii(s[2].d)];
		expect(Math.max(...innerInner)).toBeCloseTo(
			Math.max(...outerInner) + (Math.max(...outerInner) - Math.min(...outerInner)),
			0,
		);
	});

	it('labels slices with category and percentage', () => {
		const tree = render(parseChartXml(loadFixture('pie-office.xml')));
		const labels = texts(tree, 'data-label');
		expect(labels).toHaveLength(4);
		expect(labels[0]).toBe('Direct, 40%');
		expect(labels.slice(1)).toEqual(['Partner\n25%', 'Online\n20%', 'Retail\n15%']);
	});

	it('renders the pie-family fixtures', () => {
		for (const name of ['doughnut-office.xml', 'pie-title-runs.xml']) {
			expect(slices(render(parseChartXml(loadFixture(name)))).length).toBe(4);
		}
	});
});

describe('radar charts', () => {
	it('draws the web, spokes, category labels and closed series polygons', () => {
		const tree = render(parseChartXml(loadFixture('radar-office.xml')));
		expect(byRole(tree, 'gridline').length).toBeGreaterThan(0);
		expect(byRole(tree, 'axis-line').length).toBeGreaterThanOrEqual(5);
		const lines = byRole(tree, 'line') as ChartPathNode[];
		expect(lines).toHaveLength(2);
		expect(lines.every((l) => l.d.endsWith('Z'))).toBe(true);
		expect(texts(tree, 'tick-label')).toEqual(expect.arrayContaining(['1/5/2002', '1/9/2002']));
		expect(byRole(tree, 'marker')).toHaveLength(10);
	});

	it('fills filled radars', () => {
		const chart = literalChart({ type: 'radar', radarStyle: 'filled' }, [[1, 2, 3]], {
			categories: ['a', 'b', 'c'],
		});
		const tree = render(chart);
		expect(byRole(tree, 'area')).toHaveLength(1);
		expect(byRole(tree, 'marker')).toHaveLength(0);
	});
});

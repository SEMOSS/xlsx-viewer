import { describe, expect, it } from 'vitest';
import { createDefaultTheme } from '../../styles/defaults';
import type { ChartModel } from '../../types/chart';
import { assertFinite, byRole, literalChart, loadFixture, texts } from '../__tests__/helpers';
import { cachedChartData } from '../data/resolve-data';
import { parseChartXml } from '../parse/parse-chart';
import { layoutChart } from './layout-chart';
import { barGeometry } from './plot-bar';
import type { ChartRectNode, ChartRenderTree } from './types';

const theme = createDefaultTheme();
const render = (chart: ChartModel, w = 480, h = 288): ChartRenderTree => {
	const tree = layoutChart(chart, cachedChartData(chart), theme, w, h);
	assertFinite(tree);
	return tree;
};
const bars = (tree: ChartRenderTree) => byRole(tree, 'bar') as ChartRectNode[];
const cats = ['Q1', 'Q2', 'Q3', 'Q4'];

describe('bar geometry', () => {
	it('splits a category slot by gap width and overlap like Excel', () => {
		const clustered = barGeometry({ type: 'bar', series: [], axisIds: [], gapWidth: 150 }, 1, 2);
		expect(clustered.width).toBeCloseTo(1 / 3.5);
		expect(clustered.step).toBeCloseTo(clustered.width);
		expect(clustered.lead).toBeCloseTo(0.75 / 3.5);
		const overlapped = barGeometry(
			{ type: 'bar', series: [], axisIds: [], gapWidth: 219, overlap: -27 },
			1,
			2,
		);
		expect(2 * overlapped.width + overlapped.width * 0.27 + overlapped.width * 2.19).toBeCloseTo(1);
		const stacked = barGeometry(
			{ type: 'bar', series: [], axisIds: [], grouping: 'stacked', gapWidth: 100 },
			1,
			3,
		);
		expect(stacked).toEqual({ width: 0.5, step: 0, lead: 0.25 });
	});
});

describe('column and bar charts', () => {
	it('draws one bar per point with heights proportional to values', () => {
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'col' }, [[10, 20, 40, 80]], { categories: cats }),
		);
		const b = bars(tree);
		expect(b).toHaveLength(4);
		const base = b[0].y + b[0].height;
		for (const bar of b) {
			expect(bar.y + bar.height).toBeCloseTo(base, 5);
		}
		expect(b[3].height / b[0].height).toBeCloseTo(8, 5);
		expect(b[1].x).toBeGreaterThan(b[0].x);
		expect(texts(tree, 'tick-label').slice(-4)).toEqual(cats);
	});

	it('labels the value axis with Excel nice numbers', () => {
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'col' }, [[10, 87]], { categories: ['a', 'b'] }),
			480,
			480,
		);
		expect(texts(tree, 'tick-label').slice(0, 11)).toEqual([
			'0',
			'10',
			'20',
			'30',
			'40',
			'50',
			'60',
			'70',
			'80',
			'90',
			'100',
		]);
	});

	it('stacks series and normalises 100% stacks', () => {
		const stacked = render(
			literalChart(
				{ type: 'bar', barDirection: 'col', grouping: 'stacked' },
				[
					[10, 20],
					[30, 40],
				],
				{ categories: ['a', 'b'] },
			),
		);
		const [a1, b1, a2, b2] = bars(stacked);
		expect(a2.y + a2.height).toBeCloseTo(a1.y, 5);
		expect(b2.y + b2.height).toBeCloseTo(b1.y, 5);
		expect(a2.x).toBeCloseTo(a1.x, 5);
		const pct = render(
			literalChart(
				{ type: 'bar', barDirection: 'col', grouping: 'percentStacked' },
				[
					[10, 20],
					[30, 60],
				],
				{ categories: ['a', 'b'] },
			),
		);
		const p = bars(pct);
		expect(p[2].y).toBeCloseTo(p[3].y, 5);
		expect(p[0].height / (p[0].height + p[2].height)).toBeCloseTo(0.25, 5);
		const labels = texts(pct, 'tick-label');
		expect(labels).toContain('100%');
		expect(labels).not.toContain('110%');
	});

	it('stacks negative values downward from zero', () => {
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'col', grouping: 'stacked' }, [[10], [-5], [8]], {
				categories: ['a'],
			}),
		);
		const [pos1, neg, pos2] = bars(tree);
		expect(neg.y).toBeCloseTo(pos1.y + pos1.height, 5);
		expect(pos2.y + pos2.height).toBeCloseTo(pos1.y, 5);
	});

	it('draws horizontal bars with the first category at the bottom', () => {
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'bar' }, [[1, 2, 3]], {
				categories: ['first', 'second', 'third'],
			}),
		);
		const b = bars(tree);
		expect(b[0].y).toBeGreaterThan(b[2].y);
		expect(b[2].width).toBeCloseTo(3 * b[0].width, 5);
		const labels = byRole(tree, 'tick-label').filter(
			(n) => n.type === 'text' && n.lines[0] === 'first',
		);
		expect(labels).toHaveLength(1);
	});

	it('inverts negative bars when asked', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[5, -5]], {
			categories: ['a', 'b'],
		});
		chart.plots[0].series[0].invertIfNegative = true;
		const [pos, neg] = bars(render(chart));
		expect(pos.fill).toBe('#4472C4');
		expect(neg.fill).toBe('#FFFFFF');
	});

	it('uses point colors with varyColors on a single series', () => {
		const chart = literalChart(
			{ type: 'bar', barDirection: 'col', varyColors: true },
			[[1, 2, 3]],
			{ categories: ['a', 'b', 'c'] },
		);
		expect(bars(render(chart)).map((b) => b.fill)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
	});

	it('draws gradient fills', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[1]], { categories: ['a'] });
		chart.plots[0].series[0].fill = {
			type: 'gradient',
			angle: 90,
			stops: [
				{ pos: 0, color: { kind: 'srgb', value: 'FF0000' } },
				{ pos: 1, color: { kind: 'srgb', value: '0000FF' } },
			],
		};
		expect(bars(render(chart))[0].gradient).toEqual({
			kind: 'linear',
			angle: 90,
			stops: [
				{ offset: 0, color: '#FF0000' },
				{ offset: 1, color: '#0000FF' },
			],
		});
	});

	it('places value data labels outside the bar ends', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[10, 25]], {
			categories: ['a', 'b'],
		});
		chart.plots[0].series[0].dataLabels = { showValue: true };
		const tree = render(chart);
		expect(texts(tree, 'data-label')).toEqual(['10', '25']);
		const [label] = byRole(tree, 'data-label');
		const [bar] = bars(tree);
		expect(label.type === 'text' && label.y).toBeLessThan(bar.y);
	});
});

describe('category axis crowding', () => {
	it('slants and thins long labels on a crowded axis', () => {
		const labels = Array.from({ length: 40 }, (_, i) => `Category number ${i + 1}`);
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'col' }, [labels.map((_, i) => i)], {
				categories: labels,
			}),
		);
		const ticks = byRole(tree, 'tick-label').filter(
			(n) => n.type === 'text' && n.lines[0].startsWith('Category'),
		);
		expect(ticks.length).toBeLessThan(40);
		expect(ticks.every((n) => n.type === 'text' && n.rotation === -45)).toBe(true);
	});

	it('wraps multi-word labels before slanting them', () => {
		const labels = [
			'North America',
			'South America',
			'Western Europe',
			'Eastern Europe',
			'Asia Pacific',
		];
		const tree = render(
			literalChart({ type: 'bar', barDirection: 'col' }, [[1, 2, 3, 4, 5]], { categories: labels }),
			300,
			288,
		);
		const ticks = byRole(tree, 'tick-label').filter(
			(n) => n.type === 'text' && n.lines.length === 2,
		);
		expect(ticks.length).toBe(5);
	});

	it('honours an explicit rotation and multi-level categories', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[1, 2, 3, 4]], {
			categories: ['Jan', 'Feb', 'Jan', 'Feb'],
		});
		chart.axes[0].labelRotation = -90;
		chart.plots[0].series[0].categories = {
			kind: 'multiLvlStr',
			cache: ['Jan', 'Feb', 'Jan', 'Feb'],
			levels: [
				['2023', null, '2024', null],
				['Jan', 'Feb', 'Jan', 'Feb'],
			],
		};
		const tree = render(chart);
		const all = texts(tree, 'tick-label');
		expect(all).toEqual(expect.arrayContaining(['2023', '2024', 'Jan', 'Feb']));
		const jan = byRole(tree, 'tick-label').find((n) => n.type === 'text' && n.lines[0] === 'Jan');
		expect(jan?.type === 'text' && jan.rotation).toBe(-90);
	});
});

describe('line, area and combo charts', () => {
	it('breaks lines at gaps and connects over #N/A points', () => {
		const chart = literalChart({ type: 'line', marker: false }, [[1, 2, null, 4, 5]], {
			categories: ['a', 'b', 'c', 'd', 'e'],
		});
		const gap = byRole(render(chart), 'line')[0];
		expect(gap.type === 'path' && gap.d.match(/M/g)).toHaveLength(2);
		chart.plots[0].series[0].values!.cache = [1, 2, null, 4, 5];
		const spanData = cachedChartData(chart);
		spanData.plots[0].series[0].values[2] = NaN;
		const span = byRole(layoutChart(chart, spanData, theme, 480, 288), 'line')[0];
		expect(span.type === 'path' && span.d.match(/M/g)).toHaveLength(1);
	});

	it('smooths lines and cycles automatic markers', () => {
		const chart = literalChart(
			{ type: 'line', marker: true },
			[
				[1, 3, 2],
				[2, 1, 3],
			],
			{ categories: ['a', 'b', 'c'] },
		);
		chart.plots[0].series[0].smooth = true;
		const tree = render(chart);
		const [first, second] = byRole(tree, 'line');
		expect(first.type === 'path' && first.d).toContain('C');
		expect(second.type === 'path' && second.d).not.toContain('C');
		const markers = byRole(tree, 'marker');
		expect(markers.filter((m) => m.series === 0 && m.type === 'path')).toHaveLength(3);
		expect(markers.filter((m) => m.series === 1 && m.type === 'rect')).toHaveLength(3);
	});

	it('stacks lines and areas cumulatively', () => {
		const lines = render(
			literalChart(
				{ type: 'line', grouping: 'stacked', marker: false },
				[
					[1, 1],
					[2, 2],
				],
				{ categories: ['a', 'b'] },
			),
		);
		const [l1, l2] = byRole(lines, 'line').map((n) => (n.type === 'path' ? n.d : ''));
		const y = (d: string) => Number(d.split(' ')[1]);
		expect(y(l2)).toBeLessThan(y(l1));
		const areas = render(
			literalChart(
				{ type: 'area', grouping: 'stacked' },
				[
					[1, 1],
					[2, 2],
				],
				{ categories: ['a', 'b'] },
			),
		);
		expect(byRole(areas, 'area')).toHaveLength(2);
	});

	it('shares the category axis between bar and line plots and supports a secondary axis', () => {
		const combo = parseChartXml(loadFixture('combo-bar-line.xml'));
		combo.axes.push(
			{ id: 3, kind: 'cat', position: 'b', crossAxisId: 4, deleted: true },
			{ id: 4, kind: 'val', position: 'r', crossAxisId: 3, crosses: 'max' },
		);
		combo.plots[1].axisIds = [3, 4];
		combo.plots[1].series[0].values!.cache = [3000, 4100, 3800, 5200];
		const tree = render(combo);
		const labels = byRole(tree, 'tick-label').filter((n) => n.type === 'text');
		const right = labels.filter((n) => n.type === 'text' && n.anchor === 'start');
		expect(right.map((n) => (n.type === 'text' ? n.lines[0] : ''))).toContain('6000');
		expect(byRole(tree, 'bar')).toHaveLength(4);
		expect(byRole(tree, 'line')).toHaveLength(2);
	});

	it('draws reversed axes and crossing values', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[10, 20]], {
			categories: ['a', 'b'],
		});
		chart.axes[1].orientation = 'maxMin';
		const [bar] = bars(render(chart));
		expect(bar.y).toBeLessThan(60);
		const crossed = literalChart({ type: 'bar', barDirection: 'col' }, [[10, 40]], {
			categories: ['a', 'b'],
		});
		crossed.axes[0].crossesAt = 20;
		const [low, high] = bars(render(crossed));
		expect(low.y).toBeCloseTo(high.y + high.height, 5);
	});
});

describe('scatter, bubble and stock charts', () => {
	it('plots scatter points by their X values', () => {
		const chart = parseChartXml(loadFixture('scatter.xml'));
		const tree = render(chart);
		const markers = byRole(tree, 'marker').filter((m) => m.series === 0);
		expect(markers).toHaveLength(4);
		const xs = markers.map((m) => (m.type === 'circle' ? m.cx : 0));
		expect(xs[1] - xs[0]).toBeCloseTo(xs[2] - xs[1], 5);
		expect(byRole(tree, 'line')).toHaveLength(0);
	});

	it('draws trendlines', () => {
		const chart = parseChartXml(loadFixture('scatter.xml'));
		chart.plots[0].series[0].trendlines = [{ type: 'linear' }];
		const tree = render(chart);
		expect(byRole(tree, 'trendline')).toHaveLength(1);
		expect(texts(tree, 'legend-entry')).toContain('Linear (Revenue)');
	});

	it('sizes bubbles by area', () => {
		const tree = render(parseChartXml(loadFixture('bubble-office.xml')));
		const bubbles = byRole(tree, 'bubble').filter((b) => b.series === 1);
		const r = bubbles.map((b) => (b.type === 'circle' ? b.r : 0));
		expect(r[2] / r[1]).toBeCloseTo(Math.sqrt(12 / 3), 5);
	});

	it('draws high-low lines and up/down bars for stock charts', () => {
		const tree = render(parseChartXml(loadFixture('stock.xml')));
		expect(byRole(tree, 'hi-low-line')).toHaveLength(4);
		const updown = byRole(tree, 'up-down-bar') as ChartRectNode[];
		expect(updown.map((b) => b.fill)).toEqual(['#FFFFFF', '#FFFFFF', '#595959', '#FFFFFF']);
	});

	it('lays out a date axis by time', () => {
		const chart = literalChart({ type: 'line', marker: false }, [[1, 2, 3]]);
		chart.axes[0].kind = 'date';
		chart.plots[0].series[0].categories = {
			kind: 'num',
			cache: [45292, 45293, 45300],
			formatCode: 'm/d/yyyy',
		};
		const tree = render(chart);
		const d = byRole(tree, 'line')[0];
		const xs =
			d.type === 'path'
				? d.d
						.split(/[ML ]+/)
						.filter(Boolean)
						.filter((_, i) => i % 2 === 0)
						.map(Number)
				: [];
		expect((xs[2] - xs[1]) / (xs[1] - xs[0])).toBeCloseTo(7, 1);
		expect(texts(tree, 'tick-label')).toContain('1/1/2024');
	});
});

describe('axis number formats', () => {
	it('divides by display units and labels the unit', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[1200, 3400]], {
			categories: ['a', 'b'],
		});
		chart.axes[1].displayUnits = { unit: 'thousands', showLabel: true };
		const tree = render(chart);
		expect(texts(tree, 'tick-label')).toContain('4');
		expect(texts(tree, 'tick-label')).not.toContain('4000');
		const unit = byRole(tree, 'axis-title').find(
			(n) => n.type === 'text' && n.lines[0] === 'Thousands',
		);
		expect(unit?.type === 'text' && unit.rotation).toBe(-90);
	});

	it('colors tick labels from the format code', () => {
		const chart = literalChart({ type: 'bar', barDirection: 'col' }, [[-20, 30]], {
			categories: ['a', 'b'],
		});
		chart.axes[1].numFmt = { code: '0;[Red]-0', sourceLinked: false };
		const ticks = byRole(render(chart), 'tick-label').filter((n) => n.type === 'text');
		const negative = ticks.find((n) => n.type === 'text' && n.lines[0] === '-20');
		const positive = ticks.find((n) => n.type === 'text' && n.lines[0] === '20');
		expect(negative?.type === 'text' && negative.fill).toBe('#FF0000');
		expect(positive?.type === 'text' && positive.fill).toBe('#595959');
	});
});

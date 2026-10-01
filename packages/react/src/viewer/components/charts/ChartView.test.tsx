import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import {
	type ChartModel,
	type Workbook,
	cachedChartData,
	cellKey,
	createChartFromRange,
	createDefaultTheme,
	createWorkbook,
	layoutChart,
	parseChartXml,
} from 'xlsx-viewer-core';
import { ChartSvg } from './ChartSvg';
import { ChartView } from './ChartView';

function salesWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Sales'] });
	const rows: Array<Array<string | number | null>> = [
		[null, 'Revenue', 'Cost'],
		['Q1', 10, 4],
		['Q2', 20, 8],
		['Q3', 30, 12],
		['Q4', 40, 16],
	];
	rows.forEach((row, r) =>
		row.forEach((value, c) => {
			if (value !== null) {
				wb.sheets[0].cells.set(cellKey(r, c), { value });
			}
		}),
	);
	return wb;
}

const range = { top: 0, left: 0, bottom: 4, right: 2 };

describe('<ChartView>', () => {
	it('renders an accessible SVG chart', () => {
		const wb = salesWorkbook();
		const chart = createChartFromRange(wb, 0, range, 'column', { title: 'Quarterly sales' });
		render(
			<ChartView
				chart={chart}
				workbook={wb}
				sheetIndex={0}
				width={480}
				height={288}
				className='chart'
			/>,
		);
		const svg = screen.getByRole('img', { name: 'Quarterly sales, column chart with 2 series' });
		expect(svg.tagName.toLowerCase()).toBe('svg');
		expect(svg.getAttribute('viewBox')).toBe('0 0 480 288');
		expect(svg.getAttribute('class')).toBe('chart');
		expect(svg.querySelectorAll('rect[data-role="bar"]')).toHaveLength(8);
		expect(screen.getByText('Quarterly sales')).toBeTruthy();
		expect(screen.getByText('Revenue')).toBeTruthy();
		expect(screen.getByText('Cost')).toBeTruthy();
		expect(screen.getByText('Q3')).toBeTruthy();
	});

	it('re-reads cell values when the revision changes', () => {
		const wb = salesWorkbook();
		const chart = createChartFromRange(wb, 0, { top: 0, left: 0, bottom: 4, right: 1 }, 'column');
		const { container, rerender } = render(
			<ChartView
				chart={chart}
				workbook={wb}
				sheetIndex={0}
				width={480}
				height={288}
				revision={1}
			/>,
		);
		const heights = () =>
			[...container.querySelectorAll('rect[data-role="bar"]')].map((r) =>
				Number(r.getAttribute('height')),
			);
		const before = heights();
		expect(before[3] / before[0]).toBeCloseTo(4, 1);
		wb.sheets[0].cells.set(cellKey(1, 1), { value: 40 });
		rerender(
			<ChartView
				chart={chart}
				workbook={wb}
				sheetIndex={0}
				width={480}
				height={288}
				revision={1}
			/>,
		);
		expect(heights()).toEqual(before);
		rerender(
			<ChartView
				chart={chart}
				workbook={wb}
				sheetIndex={0}
				width={480}
				height={288}
				revision={2}
			/>,
		);
		const after = heights();
		expect(after[0]).toBeCloseTo(after[3], 5);
	});

	it('renders pies with a category legend', () => {
		const wb = salesWorkbook();
		const chart = createChartFromRange(wb, 0, range, 'pie');
		const { container } = render(
			<ChartView chart={chart} workbook={wb} sheetIndex={0} width={400} height={300} />,
		);
		expect(container.querySelectorAll('path[data-role="slice"]')).toHaveLength(4);
		expect(screen.getByText('Q1')).toBeTruthy();
		expect(screen.getByRole('img').getAttribute('aria-label')).toBe(
			'Revenue, pie chart with 2 series',
		);
	});

	it('renders gradients, clip paths and rich title runs as SVG definitions', () => {
		const wb = salesWorkbook();
		const chart: ChartModel = createChartFromRange(wb, 0, range, 'column');
		chart.plots[0].series[0].fill = {
			type: 'gradient',
			angle: 90,
			stops: [
				{ pos: 0, color: { kind: 'srgb', value: 'FF0000' } },
				{ pos: 1, color: { kind: 'srgb', value: '0000FF' } },
			],
		};
		chart.title = {
			text: 'Bold Red',
			paragraphs: [
				{
					runs: [
						{ text: 'Bold ', style: { bold: true } },
						{ text: 'Red', style: { color: { kind: 'srgb', value: 'FF0000' } } },
					],
				},
			],
		};
		const { container } = render(
			<ChartView chart={chart} workbook={wb} sheetIndex={0} width={480} height={288} />,
		);
		const gradient = container.querySelector('linearGradient');
		expect(gradient?.querySelectorAll('stop')).toHaveLength(2);
		const bar = container.querySelector('rect[data-role="bar"]');
		expect(bar?.getAttribute('fill')).toBe(`url(#${gradient?.getAttribute('id')})`);
		const clip = container.querySelector('clipPath');
		expect(clip).toBeTruthy();
		expect(
			container.querySelector(`g[clip-path="url(#${clip?.getAttribute('id')})"]`),
		).toBeTruthy();
		const runs = [...container.querySelectorAll('text[data-role="title"] tspan')];
		expect(runs.map((t) => t.textContent)).toEqual(['Bold ', 'Red']);
		expect(runs[0].getAttribute('font-weight')).toBe('bold');
		expect(runs[1].getAttribute('fill')).toBe('#FF0000');
	});

	it('keeps definition ids unique across charts on a page', () => {
		const wb = salesWorkbook();
		const chart = createChartFromRange(wb, 0, range, 'line');
		const { container } = render(
			<div>
				<ChartView chart={chart} workbook={wb} sheetIndex={0} width={300} height={200} />
				<ChartView chart={chart} workbook={wb} sheetIndex={0} width={300} height={200} />
			</div>,
		);
		const ids = [...container.querySelectorAll('clipPath')].map((c) => c.id);
		expect(ids.length).toBeGreaterThan(1);
		expect(new Set(ids).size).toBe(ids.length);
	});
});

describe('<ChartSvg>', () => {
	it('draws a prepared render tree, including rotated text and markers', () => {
		const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/>
			<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Visits</c:v></c:tx>
			<c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>Mon</c:v></c:pt><c:pt idx="1"><c:v>Tue</c:v></c:pt><c:pt idx="2"><c:v>Wed</c:v></c:pt></c:strLit></c:cat>
			<c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
			<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>
			<c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:title/><c:crossAx val="2"/></c:catAx>
			<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr rot="-5400000"/><a:p><a:r><a:t>Count</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="1"/></c:valAx>
			</c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`;
		const chart = parseChartXml(xml);
		const tree = layoutChart(chart, cachedChartData(chart), createDefaultTheme(), 400, 260);
		const { container } = render(<ChartSvg tree={tree} />);
		expect(container.querySelectorAll('[data-role="marker"]').length).toBeGreaterThanOrEqual(3);
		expect(container.querySelector('path[data-role="line"]')?.getAttribute('stroke')).toBe(
			'#4472C4',
		);
		const valueTitle = [...container.querySelectorAll('text[data-role="axis-title"]')].find(
			(t) => t.textContent === 'Count',
		);
		expect(valueTitle?.getAttribute('transform')).toMatch(/^rotate\(-90 /);
		expect(screen.getByText('Axis Title')).toBeTruthy();
		expect(screen.getByText('Visits')).toBeTruthy();
	});
});

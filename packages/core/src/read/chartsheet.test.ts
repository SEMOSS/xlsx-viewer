import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import type { ChartObject } from '../types/drawing';
import type { Workbook } from '../types/workbook';
import type { ReadOptions } from './context';
import { readWorkbook } from './read-workbook';
import {
	A_NS,
	C_NS,
	MAIN_NS,
	R_NS,
	XDR_NS,
	buildWorkbook,
	relsXml,
	worksheet,
} from './test-utils/package-builder';

const CHART_XML = readFileSync(join(__dirname, 'test-utils/chart-bar.xml'), 'utf8');

const CHARTSHEET =
	`<chartsheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}"><sheetPr><tabColor rgb="FF00B050"/></sheetPr>` +
	'<sheetViews><sheetView zoomScale="115" tabSelected="1" workbookViewId="0" zoomToFit="1"/></sheetViews>' +
	'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
	'<drawing r:id="rId1"/></chartsheet>';

function drawing(anchorOpen: string, anchorClose: string): string {
	return (
		`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${anchorOpen}` +
		'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
		'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>' +
		'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
		`<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame>` +
		`<xdr:clientData/>${anchorClose}</xdr:wsDr>`
	);
}

const ABSOLUTE = drawing(
	'<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="8670000" cy="6300000"/>',
	'</xdr:absoluteAnchor>',
);

async function load(drawingXml: string, options: ReadOptions = {}): Promise<Workbook> {
	return readWorkbook(
		await buildWorkbook({
			sheets: [
				{ name: 'Data', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') },
				{
					name: 'Chart1',
					kind: 'chartsheet',
					xml: CHARTSHEET,
					rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing2.xml' }],
				},
			],
			parts: {
				'xl/drawings/drawing2.xml': drawingXml,
				'xl/drawings/_rels/drawing2.xml.rels': relsXml([
					{ id: 'rIdChart', type: 'chart', target: '../charts/chart2.xml' },
				]),
				'xl/charts/chart2.xml': CHART_XML,
			},
		}),
		options,
	);
}

describe('chartsheets', () => {
	it('reads a chartsheet as one absolutely anchored chart', async () => {
		const wb = await load(ABSOLUTE);
		const sheet = wb.sheets[1];
		expect(sheet.kind).toBe('chartsheet');
		expect(sheet.tabColor).toEqual({ rgb: 'FF00B050' });
		expect(sheet.view.zoomScale).toBe(115);
		expect(sheet.pageMargins?.top).toBe(0.75);
		expect(sheet.cells.size).toBe(0);
		expect(sheet.drawings).toHaveLength(1);
		const chart = sheet.drawings[0] as ChartObject;
		expect(chart.kind).toBe('chart');
		expect(chart.chartPath).toBe('xl/charts/chart2.xml');
		expect(chart.anchor).toEqual({
			type: 'absolute',
			pos: { x: 0, y: 0 },
			ext: { cx: 8670000, cy: 6300000 },
		});
		expect(chart.raw?.name).toBe('xdr:absoluteAnchor');
		expect(sheet.source).toMatchObject({
			partPath: 'xl/chartsheets/sheet2.xml',
			drawingPath: 'xl/drawings/drawing2.xml',
		});
		expect(sheet.source?.root?.name).toBe('chartsheet');
		expect(sheet.dirty).toBe(false);
		expect(wb.warnings).toEqual([]);
	});

	it('re-anchors a chart that the drawing places with cell anchors', async () => {
		const wb = await load(
			drawing(
				'<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
					'<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
				'</xdr:twoCellAnchor>',
			),
		);
		const anchor = wb.sheets[1].drawings[0].anchor;
		expect(anchor.type).toBe('absolute');
		expect(anchor.type === 'absolute' && anchor.pos).toEqual({ x: 0, y: 0 });
		expect(anchor.type === 'absolute' && anchor.ext.cx).toBeGreaterThan(0);
	});

	it('keeps the drawing path but skips the chart when drawings are skipped', async () => {
		const wb = await load(ABSOLUTE, { skipDrawings: true });
		expect(wb.sheets[1].drawings).toEqual([]);
		expect(wb.sheets[1].source?.drawingPath).toBe('xl/drawings/drawing2.xml');
	});

	it('warns when the chartsheet drawing has no chart', async () => {
		const wb = await load(`<xdr:wsDr xmlns:xdr="${XDR_NS}"/>`);
		expect(wb.sheets[1].drawings).toEqual([]);
		expect(wb.warnings.map((w) => w.code)).toEqual(['unsupported-chart']);
	});
});

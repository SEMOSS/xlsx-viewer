import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import type {
	ChartObject,
	DrawingObject,
	GroupObject,
	PictureObject,
	ShapeObject,
	UnknownDrawingObject,
} from '../types/drawing';
import type { Workbook } from '../types/workbook';
import type { ReadOptions } from './context';
import { readWorkbook } from './read-workbook';
import {
	A_NS,
	C_NS,
	MAIN_NS,
	PNG_1X1,
	R_NS,
	XDR_NS,
	buildWorkbook,
	worksheet,
} from './test-utils/package-builder';

const CHART_XML = readFileSync(join(__dirname, 'test-utils/chart-bar.xml'), 'utf8');
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function from(col: number, row: number): string {
	return `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>`;
}

function to(col: number, row: number): string {
	return `<xdr:to><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;
}

const PIC =
	'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="Logo"><a:hlinkClick xmlns:r="' +
	R_NS +
	'" r:id="rIdLink"/></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
	'<xdr:blipFill><a:blip xmlns:r="' +
	R_NS +
	'" r:embed="rIdImg"/><a:srcRect l="1000" t="2000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
	'<xdr:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
	'<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></xdr:spPr></xdr:pic>';

const SHAPE =
	'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Rounded 2" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
	'<xdr:spPr><a:xfrm flipV="1"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' +
	'<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/><a:lumOff val="25%"/></a:schemeClr></a:solidFill>' +
	'<a:ln w="28575" cap="rnd"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill><a:prstDash val="dash"/>' +
	'<a:headEnd type="oval"/><a:tailEnd type="triangle" w="lg" len="sm"/></a:ln></xdr:spPr>' +
	'<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr" vert="vert270"/><a:lstStyle/>' +
	'<a:p><a:pPr algn="ctr"><a:defRPr sz="1000"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1" i="1" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>' +
	'<a:latin typeface="Arial"/></a:rPr><a:t>Hello</a:t></a:r><a:br><a:rPr sz="1400"/></a:br><a:r><a:rPr b="0"/><a:t>World</a:t></a:r></a:p>' +
	'<a:p><a:endParaRPr lang="en-US" sz="1200"/></a:p></xdr:txBody></xdr:sp>';

const STYLED_TEXTBOX =
	'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="4" name="TextBox 3"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>' +
	'<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
	'<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>' +
	'<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
	'<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>' +
	'<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Styled</a:t></a:r></a:p></xdr:txBody></xdr:sp>';

const CONNECTOR =
	'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="Connector 4"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>' +
	'<xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr>' +
	'<xdr:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
	'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style></xdr:cxnSp>';

const GROUP =
	'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="6" name="Group 5"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>' +
	'<xdr:grpSpPr><a:xfrm rot="600000"><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/><a:chOff x="10" y="20"/><a:chExt cx="300" cy="400"/></a:xfrm></xdr:grpSpPr>' +
	'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7" name="Oval 6"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="150" cy="200"/></a:xfrm>' +
	'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
	'<a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></xdr:spPr></xdr:sp>' +
	'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="Rect 7"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="160" y="20"/><a:ext cx="150" cy="200"/></a:xfrm>' +
	'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:pattFill prst="dkDnDiag"><a:fgClr><a:prstClr val="black"/></a:fgClr><a:bgClr><a:hslClr hue="0" sat="0%" lum="100%"/></a:bgClr></a:pattFill></xdr:spPr></xdr:sp>' +
	'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9" name="Filled 8"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="10" y="230"/><a:ext cx="20" cy="20"/></a:xfrm>' +
	'<a:custGeom><a:pathLst/></a:custGeom><a:blipFill><a:blip xmlns:r="' +
	R_NS +
	'" r:embed="rIdImg"/><a:stretch/></a:blipFill><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp>' +
	'</xdr:grpSp>';

const CHART_FRAME = (id: number, name: string): string =>
	`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
	'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
	`<a:graphic><a:graphicData uri="${CHART_URI}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame>`;

const CHARTEX_FRAME =
	'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="20" name="Chart 19"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
	'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">' +
	'<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="' +
	R_NS +
	'" r:id="rIdChartEx"/></a:graphicData></a:graphic></xdr:graphicFrame>';

const DRAWING =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:mc="${MC_NS}">` +
	`<xdr:twoCellAnchor editAs="oneCell">${from(1, 2)}${to(4, 8)}${PIC}<xdr:clientData/></xdr:twoCellAnchor>` +
	`<xdr:oneCellAnchor>${from(0, 0)}<xdr:ext cx="914400" cy="457200"/>${SHAPE}<xdr:clientData/></xdr:oneCellAnchor>` +
	`<xdr:twoCellAnchor>${from(5, 5)}${to(7, 7)}${STYLED_TEXTBOX}<xdr:clientData/></xdr:twoCellAnchor>` +
	`<xdr:absoluteAnchor><xdr:pos x="100" y="200"/><xdr:ext cx="300" cy="0"/>${CONNECTOR}<xdr:clientData/></xdr:absoluteAnchor>` +
	`<xdr:twoCellAnchor>${from(0, 10)}${to(3, 14)}${GROUP}<xdr:clientData/></xdr:twoCellAnchor>` +
	`<xdr:twoCellAnchor>${from(8, 0)}${to(15, 15)}${CHART_FRAME(10, 'Chart 9')}<xdr:clientData/></xdr:twoCellAnchor>` +
	`<mc:AlternateContent><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">` +
	`<xdr:twoCellAnchor>${from(0, 20)}${to(5, 30)}${CHARTEX_FRAME}<xdr:clientData/></xdr:twoCellAnchor></mc:Choice>` +
	`<mc:Fallback><xdr:twoCellAnchor>${from(0, 20)}${to(5, 30)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="21" name="Fallback"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
	'<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>This chart isn\'t available in your version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp>' +
	'<xdr:clientData/></xdr:twoCellAnchor></mc:Fallback></mc:AlternateContent>' +
	`<xdr:twoCellAnchor>${from(9, 20)}${to(12, 30)}<mc:AlternateContent><mc:Choice Requires="a14">${CHART_FRAME(30, 'Chart 29')}</mc:Choice>` +
	'<mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="31" name="Old"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>' +
	`<xdr:twoCellAnchor>${from(0, 40)}${to(4, 44)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Diagram 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
	'<xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
	`<xdr:oneCellAnchor>${from(0, 50)}<xdr:ext cx="10" cy="10"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="40" name="Gone"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
	`<xdr:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rIdMissing"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>` +
	'</xdr:wsDr>';

const DRAWING_RELS = [
	{ id: 'rIdImg', type: 'image', target: '../media/image1.png' },
	{ id: 'rIdLink', type: 'hyperlink', target: 'https://example.com/', external: true },
	{ id: 'rIdChart', type: 'chart', target: '../charts/chart1.xml' },
	{
		id: 'rIdChartEx',
		type: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
		target: '../charts/chartEx1.xml',
	},
	{ id: 'rIdMissing', type: 'image', target: '../media/nope.png' },
];

async function load(options: ReadOptions = {}): Promise<Workbook> {
	const { relsXml } = await import('./test-utils/package-builder');
	const bytes = await buildWorkbook({
		sheets: [
			{
				name: 'Pics',
				xml: worksheet('', { after: '<drawing r:id="rIdD"/>' }),
				rels: [{ id: 'rIdD', type: 'drawing', target: '../drawings/drawing1.xml' }],
			},
		],
		parts: {
			'xl/drawings/drawing1.xml': DRAWING,
			'xl/drawings/_rels/drawing1.xml.rels': relsXml(DRAWING_RELS),
			'xl/media/image1.png': PNG_1X1,
			'xl/charts/chart1.xml': CHART_XML,
		},
	});
	return readWorkbook(bytes, options);
}

function byName(objects: DrawingObject[], name: string): DrawingObject {
	const o = objects.find((x) => x.name === name);
	expect(o, name).toBeDefined();
	return o!;
}

describe('drawings', () => {
	it('reads every anchored object in order', async () => {
		const wb = await load();
		const d = wb.sheets[0].drawings;
		expect(d.map((o) => [o.kind, o.name])).toEqual([
			['picture', 'Picture 1'],
			['shape', 'Rounded 2'],
			['shape', 'TextBox 3'],
			['shape', 'Connector 4'],
			['group', 'Group 5'],
			['chart', 'Chart 9'],
			['shape', 'Fallback'],
			['chart', 'Chart 29'],
			['unknown', 'Diagram 1'],
			['unknown', 'Gone'],
		]);
		expect(wb.sheets[0].source?.drawingPath).toBe('xl/drawings/drawing1.xml');
		for (const o of d) {
			expect(o.dirty).toBe(false);
		}
	});

	it('reads pictures with crop, rotation, outline, hyperlink and media', async () => {
		const wb = await load();
		const pic = byName(wb.sheets[0].drawings, 'Picture 1') as PictureObject;
		expect(pic).toMatchObject({
			id: '2',
			kind: 'picture',
			description: 'Logo',
			mediaPath: 'xl/media/image1.png',
			crop: { l: 1000, t: 2000 },
			rotation: 90,
			flipH: true,
			hyperlink: 'https://example.com/',
			line: { width: 12700, fill: { type: 'solid', color: { kind: 'srgb', value: 'FF0000' } } },
			anchor: {
				type: 'twoCell',
				editAs: 'oneCell',
				from: { col: 1, colOff: 9525, row: 2, rowOff: 19050 },
				to: { col: 4, colOff: 0, row: 8, rowOff: 0 },
			},
		});
		expect(pic.raw?.name).toBe('xdr:twoCellAnchor');
		const media = wb.media.get('xl/media/image1.png');
		expect(media?.contentType).toBe('image/png');
		expect(media?.data).toEqual(PNG_1X1);
		expect(wb.media.size).toBe(1);
	});

	it('reads preset shapes with fills, outlines, transforms and rich text', async () => {
		const wb = await load();
		const shape = byName(wb.sheets[0].drawings, 'Rounded 2') as ShapeObject;
		expect(shape.anchor).toEqual({
			type: 'oneCell',
			from: { col: 0, colOff: 9525, row: 0, rowOff: 19050 },
			ext: { cx: 914400, cy: 457200 },
		});
		expect(shape).toMatchObject({
			id: '3',
			hidden: true,
			flipV: true,
			geometry: 'roundRect',
			adjust: { adj: 16667 },
			fill: {
				type: 'solid',
				color: {
					kind: 'scheme',
					value: 'accent1',
					transforms: [
						{ name: 'lumMod', val: 75000 },
						{ name: 'lumOff', val: 25000 },
					],
				},
			},
			line: {
				width: 28575,
				cap: 'rnd',
				dash: 'dash',
				fill: { type: 'solid', color: { kind: 'sys', value: '000000' } },
				headEnd: { type: 'oval' },
				tailEnd: { type: 'triangle', w: 'lg', len: 'sm' },
			},
		});
		expect(shape.text).toEqual({
			anchor: 'ctr',
			insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
			wrap: false,
			vertical: 'vert270',
			paragraphs: [
				{
					align: 'ctr',
					runs: [
						{
							text: 'Hello',
							style: {
								size: 14,
								bold: true,
								italic: true,
								underline: true,
								strike: true,
								color: { kind: 'srgb', value: '00FF00' },
								font: 'Arial',
							},
						},
						{ text: '\n', style: { size: 14 } },
						{ text: 'World', style: { size: 10, bold: false } },
					],
				},
				{ runs: [{ text: '', style: { size: 12 } }] },
			],
		});
	});

	it('uses the shape style references as fallback fill, line and text colors', async () => {
		const wb = await load();
		const box = byName(wb.sheets[0].drawings, 'TextBox 3') as ShapeObject;
		expect(box.isTextBox).toBe(true);
		expect(box.fill).toEqual({ type: 'solid', color: { kind: 'scheme', value: 'accent1' } });
		expect(box.line).toEqual({
			width: 12700,
			fill: {
				type: 'solid',
				color: { kind: 'scheme', value: 'accent1', transforms: [{ name: 'shade', val: 50000 }] },
			},
		});
		expect(box.text?.paragraphs[0].runs[0]).toEqual({
			text: 'Styled',
			style: { color: { kind: 'scheme', value: 'lt1' } },
		});
		const connector = byName(wb.sheets[0].drawings, 'Connector 4') as ShapeObject;
		expect(connector).toMatchObject({
			isConnector: true,
			geometry: 'straightConnector1',
			anchor: { type: 'absolute', pos: { x: 100, y: 200 }, ext: { cx: 300, cy: 0 } },
			line: {
				tailEnd: { type: 'triangle' },
				width: 6350,
				fill: { type: 'solid', color: { kind: 'scheme', value: 'accent1' } },
			},
		});
		expect(connector.fill).toBeUndefined();
	});

	it('reads groups with child coordinates and nested fills', async () => {
		const wb = await load();
		const group = byName(wb.sheets[0].drawings, 'Group 5') as GroupObject;
		expect(group.rotation).toBe(10);
		expect(group.childOffset).toEqual({ x: 10, y: 20 });
		expect(group.childExtent).toEqual({ cx: 300, cy: 400 });
		expect(group.children.map((c) => c.name)).toEqual(['Oval 6', 'Rect 7', 'Filled 8']);
		const [oval, rect, filled] = group.children as ShapeObject[];
		expect(oval.anchor).toEqual({
			type: 'absolute',
			pos: { x: 10, y: 20 },
			ext: { cx: 150, cy: 200 },
		});
		expect(oval.raw?.name).toBe('xdr:sp');
		expect(oval.fill).toEqual({
			type: 'gradient',
			angle: 90,
			stops: [
				{ pos: 0, color: { kind: 'srgb', value: 'FFFFFF' } },
				{ pos: 1, color: { kind: 'scheme', value: 'accent2' } },
			],
		});
		expect(rect.fill).toEqual({
			type: 'pattern',
			preset: 'dkDnDiag',
			fg: { kind: 'preset', value: 'black' },
			bg: { kind: 'hsl', value: '0,0,100000' },
		});
		expect(filled.geometry).toBe('custom');
		expect(filled.fill).toEqual({ type: 'picture', mediaPath: 'xl/media/image1.png' });
		expect(filled.line).toEqual({ fill: { type: 'none' } });
	});

	it('loads chart parts and picks understood alternate content', async () => {
		const wb = await load();
		const d = wb.sheets[0].drawings;
		const chart = byName(d, 'Chart 9') as ChartObject;
		expect(chart.kind).toBe('chart');
		expect(chart.chartPath).toBe('xl/charts/chart1.xml');
		expect(Array.isArray(chart.chart.plots)).toBe(true);
		expect(chart.anchor).toMatchObject({ type: 'twoCell', to: { col: 15, row: 15 } });
		const wrapped = byName(d, 'Chart 29') as ChartObject;
		expect(wrapped.kind).toBe('chart');
		expect(wrapped.raw?.name).toBe('xdr:twoCellAnchor');
		const fallback = byName(d, 'Fallback') as ShapeObject;
		expect(fallback.raw?.name).toBe('mc:AlternateContent');
		expect(fallback.text?.paragraphs[0].runs[0].text).toContain("isn't available");
		expect((byName(d, 'Diagram 1') as UnknownDrawingObject).label).toBe('SmartArt');
		expect((byName(d, 'Gone') as UnknownDrawingObject).label).toBe('Picture (image missing)');
	});

	it('keeps object ids unique and reports unsupported content', async () => {
		const wb = await load();
		const ids = wb.sheets[0].drawings.map((o) => o.id);
		expect(new Set(ids).size).toBe(ids.length);
		expect(ids[0]).toBe('2');
		expect(byName(wb.sheets[0].drawings, 'Diagram 1').id).not.toBe('2');
		const codes = wb.warnings.map((w) => w.code).sort();
		expect(codes).toEqual(['missing-part', 'unsupported-chart']);
	});

	it('skips drawings when asked but keeps the drawing path', async () => {
		const wb = await load({ skipDrawings: true });
		expect(wb.sheets[0].drawings).toEqual([]);
		expect(wb.sheets[0].source?.drawingPath).toBe('xl/drawings/drawing1.xml');
		expect(wb.media.size).toBe(0);
	});
});

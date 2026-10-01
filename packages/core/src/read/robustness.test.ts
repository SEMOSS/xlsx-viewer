import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import type { PictureObject, UnknownDrawingObject } from '../types/drawing';
import { serializeElement } from '../xml/serializer';
import { parseXml } from '../xml/parser';
import { decodeXmlBytes } from './part-io';
import { readWorkbook } from './read-workbook';
import { normalizeSpreadsheetTree, scanRootTag, spreadsheetPrefix } from './xml-normalize';
import {
	A_NS,
	MAIN_NS,
	PNG_1X1,
	R_NS,
	X14_NS,
	XDR_NS,
	buildWorkbook,
	relsXml,
	workbookParts,
	worksheet,
	zipParts,
} from './test-utils/package-builder';

function utf16le(text: string): Uint8Array {
	const out = new Uint8Array(2 + text.length * 2);
	out[0] = 0xff;
	out[1] = 0xfe;
	for (let i = 0; i < text.length; i++) {
		out[2 + i * 2] = text.charCodeAt(i) & 0xff;
		out[3 + i * 2] = text.charCodeAt(i) >> 8;
	}
	return out;
}

describe('namespace normalisation', () => {
	it('locates the root tag after the prolog, comments and a BOM', () => {
		const xml =
			'﻿<?xml version="1.0"?>\n<!-- hi <x> -->\n<x:worksheet xmlns:x="' + MAIN_NS + "\" a='1>2'>";
		const tag = scanRootTag(xml);
		expect(tag?.name).toBe('x:worksheet');
		expect(tag?.attrs.a).toBe('1>2');
		expect(tag && xml.slice(tag.end)).toBe('');
		expect(spreadsheetPrefix(tag!.attrs)).toBe('x');
	});

	it('rewrites prefixed trees into equivalent default-namespace XML', () => {
		const doc = parseXml(
			`<s:worksheet xmlns:s="${MAIN_NS}" xmlns:rel="${R_NS}"><s:drawing rel:id="rId1"/><s:extLst><s:ext><x14:id xmlns:x14="${X14_NS}">1</x14:id></s:ext></s:extLst></s:worksheet>`,
		);
		expect(normalizeSpreadsheetTree(doc.root)).toBe('s');
		expect(serializeElement(doc.root)).toBe(
			`<worksheet xmlns="${MAIN_NS}" xmlns:s="${MAIN_NS}" xmlns:rel="${R_NS}" xmlns:r="${R_NS}">` +
				`<drawing r:id="rId1"/><extLst><ext><x14:id xmlns:x14="${X14_NS}">1</x14:id></ext></extLst></worksheet>`,
		);
	});

	it('leaves standard documents untouched', () => {
		const doc = parseXml(
			`<worksheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}"><drawing r:id="rId1"/></worksheet>`,
		);
		const before = serializeElement(doc.root);
		expect(normalizeSpreadsheetTree(doc.root)).toBe('');
		expect(serializeElement(doc.root)).toBe(before);
	});
});

describe('part decoding', () => {
	it('decodes UTF-8 (with or without BOM) and UTF-16 parts', () => {
		expect(decodeXmlBytes(new TextEncoder().encode('﻿<a>é</a>'))).toBe('<a>é</a>');
		expect(decodeXmlBytes(utf16le('<a>ü</a>'))).toBe('<a>ü</a>');
	});

	it('reads a workbook whose sheet part is UTF-16 encoded', async () => {
		const parts = workbookParts({ sheets: [{ name: 'U', xml: '' }] });
		parts['xl/worksheets/sheet1.xml'] = utf16le(
			'<?xml version="1.0" encoding="UTF-16"?>' +
				worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>ünï</t></is></c></row>').replace(
					/^<\?xml[^>]*\?>\n/,
					'',
				),
		);
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets[0].cells.get(cellKey(0, 0))?.value).toBe('ünï');
	});
});

describe('odd but valid packages', () => {
	it('reads dialogsheets and macrosheets with their kinds', async () => {
		const parts = workbookParts({
			sheets: [
				{
					name: 'Dlg',
					kind: 'dialogsheet',
					xml: `<dialogsheet xmlns="${MAIN_NS}"><sheetViews><sheetView zoomScale="90" workbookViewId="0"/></sheetViews></dialogsheet>`,
				},
			],
			workbookRels: [
				{
					id: 'rIdM',
					type: 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
					target: 'macrosheets/sheet1.xml',
				},
			],
		});
		parts['xl/workbook.xml'] = parts['xl/workbook.xml']
			.toString()
			.replace('</sheets>', '<sheet name="Macro1" sheetId="9" r:id="rIdM"/></sheets>');
		parts['xl/macrosheets/sheet1.xml'] =
			`<xm:macrosheet xmlns="${MAIN_NS}" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
			'<sheetData><row r="1"><c r="A1" t="str"><f>GET.CELL(1)</f><v>x</v></c></row></sheetData></xm:macrosheet>';
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets.map((s) => [s.name, s.kind])).toEqual([
			['Dlg', 'dialogsheet'],
			['Macro1', 'macrosheet'],
		]);
		expect(wb.sheets[0].view.zoomScale).toBe(90);
		expect(wb.sheets[1].cells.get(cellKey(0, 0))).toEqual({ value: 'x', formula: 'GET.CELL(1)' });
	});

	it('resolves percent-encoded relationship targets', async () => {
		const drawing =
			`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
			'<xdr:ext cx="10" cy="10"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
			`<xdr:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'S',
						xml: worksheet('', { after: '<drawing r:id="rIdD"/>' }),
						rels: [{ id: 'rIdD', type: 'drawing', target: '../drawings/drawing1.xml' }],
					},
				],
				parts: {
					'xl/drawings/drawing1.xml': drawing,
					'xl/drawings/_rels/drawing1.xml.rels': relsXml([
						{ id: 'rId1', type: 'image', target: '../media/my%20image.png' },
					]),
					'xl/media/my image.png': PNG_1X1,
				},
			}),
		);
		const pic = wb.sheets[0].drawings[0] as PictureObject;
		expect(pic.mediaPath).toBe('xl/media/my image.png');
		expect(wb.media.get('xl/media/my image.png')?.contentType).toBe('image/png');
	});

	it('survives a malformed workbook relationships part with warnings', async () => {
		const parts = workbookParts({ sheets: [{ name: 'A', xml: worksheet('') }] });
		parts['xl/_rels/workbook.xml.rels'] = '<Relationships><Relationship';
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets.map((s) => s.name)).toEqual(['A']);
		expect(wb.warnings.map((w) => w.code)).toEqual(['unreadable-part', 'missing-part']);
	});

	it('keeps cells when the drawing part is broken and reports missing chart parts', async () => {
		const broken = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'S',
						xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>', {
							after: '<drawing r:id="rIdD"/>',
						}),
						rels: [{ id: 'rIdD', type: 'drawing', target: '../drawings/drawing1.xml' }],
					},
				],
				parts: { 'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:oneCellAnchor>' },
			}),
		);
		expect(broken.sheets[0].cells.size).toBe(1);
		expect(broken.sheets[0].drawings).toEqual([]);
		expect(broken.warnings.map((w) => [w.code, w.part])).toEqual([
			['unreadable-part', 'xl/drawings/drawing1.xml'],
		]);

		const frame =
			`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/>` +
			'<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Lost"/></xdr:nvGraphicFramePr><a:graphic>' +
			'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
			`xmlns:r="${R_NS}" r:id="rIdX"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`;
		const missing = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'S',
						xml: worksheet('', { after: '<drawing r:id="rIdD"/>' }),
						rels: [{ id: 'rIdD', type: 'drawing', target: '../drawings/drawing1.xml' }],
					},
				],
				parts: { 'xl/drawings/drawing1.xml': frame },
			}),
		);
		expect((missing.sheets[0].drawings[0] as UnknownDrawingObject).label).toBe(
			'Chart (missing part)',
		);
		expect(missing.warnings.map((w) => w.code)).toEqual(['missing-part']);
	});

	it('reads ISO date cells in a 1904 workbook and CDATA values', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'D',
						xml: worksheet(
							'<row r="1"><c r="A1" t="d"><v>1904-01-03T12:00:00</v></c><c r="B1"><v><![CDATA[12.5]]></v></c>' +
								'<c r="C1" t="str"><v><![CDATA[<b>&amp;]]></v></c></row>',
						),
					},
				],
				workbookBeforeSheets: '<workbookPr date1904="1"/>',
			}),
		);
		const cells = wb.sheets[0].cells;
		expect(cells.get(cellKey(0, 0))?.value).toBe(2.5);
		expect(cells.get(cellKey(0, 1))?.value).toBe(12.5);
		expect(cells.get(cellKey(0, 2))?.value).toBe('<b>&amp;');
	});

	it('normalises overlapping column records', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'C',
						xml: worksheet('', {
							before:
								'<cols><col min="1" max="5" width="10"/><col min="3" max="8" width="20"/><col min="4" max="4" width="30"/></cols>',
						}),
					},
				],
			}),
		);
		expect(wb.sheets[0].cols).toEqual([
			{ min: 0, max: 4, width: 10 },
			{ min: 5, max: 7, width: 20 },
		]);
	});

	it('warns about slicers, timelines and query tables', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'Q',
						xml: worksheet('', {
							after:
								'<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>' +
								`<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14_NS}"><x14:slicerList><x14:slicer r:id="rIdS"/></x14:slicerList></ext></extLst>`,
						}),
						rels: [{ id: 'rIdT', type: 'table', target: '../tables/table1.xml' }],
					},
				],
				workbookRels: [
					{
						id: 'rIdSC',
						type: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
						target: 'slicerCaches/slicerCache1.xml',
					},
				],
				parts: {
					'xl/tables/table1.xml': `<table xmlns="${MAIN_NS}" id="1" name="Q1" displayName="Q1" ref="A1:B2" tableType="queryTable"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>`,
				},
			}),
		);
		expect(wb.warnings.map((w) => w.code).sort()).toEqual(['external-data', 'slicer', 'slicer']);
		expect(wb.sheets[0].tables[0].name).toBe('Q1');
	});
});

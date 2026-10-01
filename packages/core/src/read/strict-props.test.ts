import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import type { Workbook } from '../types/workbook';
import { readWorkbook } from './read-workbook';
import {
	MAIN_NS,
	R_NS,
	buildWorkbook,
	relsXml,
	worksheet,
	zipParts,
} from './test-utils/package-builder';

const STRICT_MAIN = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const STRICT_R = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

function v(wb: Workbook, sheet: number, row: number, col: number): unknown {
	return wb.sheets[sheet].cells.get(cellKey(row, col))?.value;
}

describe('Strict OOXML', () => {
	it('reads a Strict package (purl.oclc.org namespaces, ISO date cells)', async () => {
		const parts = {
			'[Content_Types].xml':
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
				'<Default Extension="xml" ContentType="application/xml"/></Types>',
			'_rels/.rels': relsXml([
				{ id: 'rId1', type: `${STRICT_R}/officeDocument`, target: 'xl/workbook.xml' },
			]),
			'xl/workbook.xml':
				`<workbook xmlns="${STRICT_MAIN}" xmlns:r="${STRICT_R}" conformance="strict"><workbookPr date1904="false"/>` +
				'<bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Strict Sheet" sheetId="1" r:id="rId1"/></sheets>' +
				'<definedNames><definedName name="Total">\'Strict Sheet\'!$C$1</definedName></definedNames></workbook>',
			'xl/_rels/workbook.xml.rels': relsXml([
				{ id: 'rId1', type: `${STRICT_R}/worksheet`, target: 'worksheets/sheet1.xml' },
				{ id: 'rId2', type: `${STRICT_R}/sharedStrings`, target: 'sharedStrings.xml' },
				{ id: 'rId3', type: `${STRICT_R}/styles`, target: 'styles.xml' },
			]),
			'xl/sharedStrings.xml': `<sst xmlns="${STRICT_MAIN}"><si><t>strict text</t></si></sst>`,
			'xl/styles.xml':
				`<styleSheet xmlns="${STRICT_MAIN}"><fonts><font><b val="true"/><sz val="12"/></font></fonts>` +
				'<cellXfs><xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="true"/></cellXfs></styleSheet>',
			'xl/worksheets/sheet1.xml':
				`<worksheet xmlns="${STRICT_MAIN}" xmlns:r="${STRICT_R}"><sheetViews><sheetView tabSelected="true" workbookViewId="0"/></sheetViews>` +
				'<sheetData><row r="1" customHeight="true" ht="20"><c r="A1" t="d" s="0"><v>2024-03-01T00:00:00</v></c><c r="B1" t="s"><v>0</v></c>' +
				'<c r="C1"><f>A1+1</f><v>45353</v></c><c r="D1" t="b"><v>true</v></c></row></sheetData>' +
				'<hyperlinks><hyperlink ref="B1" r:id="rIdH"/></hyperlinks></worksheet>',
			'xl/worksheets/_rels/sheet1.xml.rels': relsXml([
				{
					id: 'rIdH',
					type: `${STRICT_R}/hyperlink`,
					target: 'https://strict.example/',
					external: true,
				},
			]),
		};
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets[0].name).toBe('Strict Sheet');
		expect(v(wb, 0, 0, 0)).toBe(45352);
		expect(v(wb, 0, 0, 1)).toBe('strict text');
		expect(wb.sheets[0].cells.get(cellKey(0, 2))).toEqual({ value: 45353, formula: 'A1+1' });
		expect(v(wb, 0, 0, 3)).toBe(true);
		expect(wb.sheets[0].rows.get(0)).toEqual({ height: 20, customHeight: true });
		expect(wb.sheets[0].view.tabSelected).toBe(true);
		expect(wb.sheets[0].hyperlinks[0].target).toBe('https://strict.example/');
		expect(wb.styles.fonts[0]).toEqual({ bold: true, size: 12 });
		expect(wb.styles.cellXfs[0].applyNumberFormat).toBe(true);
		expect(wb.definedNames).toEqual([{ name: 'Total', formula: "'Strict Sheet'!$C$1" }]);
		expect(wb.date1904).toBe(false);
		expect(wb.warnings).toEqual([]);
	});
});

describe('prefixed SpreadsheetML', () => {
	it('reads parts whose main-namespace elements use a prefix and a renamed relationships prefix', async () => {
		const parts = {
			'[Content_Types].xml':
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
			'_rels/.rels': relsXml([{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' }]),
			'xl/workbook.xml':
				`<x:workbook xmlns:x="${MAIN_NS}" xmlns:rel="${R_NS}"><x:bookViews><x:workbookView activeTab="0"/></x:bookViews>` +
				'<x:sheets><x:sheet name="Pfx" sheetId="7" rel:id="rId1"/></x:sheets><x:calcPr calcId="1"/></x:workbook>',
			'xl/_rels/workbook.xml.rels': relsXml([
				{ id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
				{ id: 'rId2', type: 'sharedStrings', target: 'sharedStrings.xml' },
			]),
			'xl/sharedStrings.xml': `<x:sst xmlns:x="${MAIN_NS}"><x:si><x:t>prefixed</x:t></x:si></x:sst>`,
			'xl/worksheets/sheet1.xml':
				`<x:worksheet xmlns:x="${MAIN_NS}" xmlns:rel="${R_NS}"><x:sheetViews><x:sheetView zoomScale="150" workbookViewId="0"/></x:sheetViews>` +
				'<x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:v>2</x:v></x:c></x:row></x:sheetData>' +
				'<x:mergeCells count="1"><x:mergeCell ref="C1:D1"/></x:mergeCells><x:hyperlinks><x:hyperlink ref="A1" rel:id="rIdH"/></x:hyperlinks></x:worksheet>',
			'xl/worksheets/_rels/sheet1.xml.rels': relsXml([
				{ id: 'rIdH', type: 'hyperlink', target: 'https://prefixed.example/', external: true },
			]),
		};
		const wb = await readWorkbook(await zipParts(parts));
		const sheet = wb.sheets[0];
		expect(sheet.name).toBe('Pfx');
		expect(sheet.sheetId).toBe(7);
		expect(v(wb, 0, 0, 0)).toBe('prefixed');
		expect(v(wb, 0, 0, 1)).toBe(2);
		expect(sheet.view.zoomScale).toBe(150);
		expect(sheet.merges).toEqual([{ top: 0, left: 2, bottom: 0, right: 3 }]);
		expect(sheet.hyperlinks[0].target).toBe('https://prefixed.example/');
		const root = sheet.source!.root!;
		expect(root.name).toBe('worksheet');
		expect(root.attrs.xmlns).toBe(MAIN_NS);
		expect(root.attrs['xmlns:r']).toBe(R_NS);
		expect(wb.calc.calcId).toBe(1);
	});
});

describe('dynamic arrays through xl/metadata.xml', () => {
	it('marks formulas whose cell metadata is XLDAPR as dynamic', async () => {
		const metadata =
			`<metadata xmlns="${MAIN_NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
			'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" ' +
			'splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
			'<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
			'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
			'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'D',
						xml: worksheet(
							'<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn._xlws.SORT(B1:B3)</f><v>1</v></c><c r="B1"><v>3</v></c></row>' +
								'<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>2</v></c></row>' +
								'<row r="5"><c r="A5"><f t="array" ref="A5">SUM(B1:B3*2)</f><v>12</v></c></row>',
						),
					},
				],
				workbookRels: [{ id: 'rIdM', type: 'sheetMetadata', target: 'metadata.xml' }],
				parts: { 'xl/metadata.xml': metadata },
			}),
		);
		const cells = wb.sheets[0].cells;
		expect(cells.get(cellKey(0, 0))).toEqual({
			value: 1,
			formula: 'SORT(B1:B3)',
			array: { ref: { top: 0, left: 0, bottom: 2, right: 0 }, dynamic: true },
			cm: 1,
		});
		expect(cells.get(cellKey(1, 0))?.arrayAnchor).toBe(cellKey(0, 0));
		expect(cells.get(cellKey(2, 0))?.arrayAnchor).toBe(cellKey(0, 0));
		expect(cells.get(cellKey(1, 1))?.arrayAnchor).toBeUndefined();
		expect(cells.get(cellKey(4, 0))?.array).toEqual({
			ref: { top: 4, left: 0, bottom: 4, right: 0 },
			dynamic: false,
		});
	});
});

describe('document properties', () => {
	it('reads core, extended and custom properties', async () => {
		const core =
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
			'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
			'<dc:title>Budget</dc:title><dc:subject>Finance</dc:subject><dc:creator>Ann</dc:creator><cp:keywords>q1, plan</cp:keywords>' +
			'<dc:description>Quarterly budget</dc:description><cp:lastModifiedBy>Ben</cp:lastModifiedBy><cp:category>Reports</cp:category>' +
			'<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-03T04:05:06Z</dcterms:modified>' +
			'</cp:coreProperties>';
		const app =
			'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
			'<Application>Microsoft Excel</Application><Manager>Mia</Manager><Company>ACME &amp; Co</Company><AppVersion>16.0300</AppVersion></Properties>';
		const custom =
			'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
			'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Project"><vt:lpwstr>Apollo</vt:lpwstr></property>' +
			'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Version"><vt:i4>42</vt:i4></property>' +
			'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Approved"><vt:bool>true</vt:bool></property>' +
			'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Ratio"><vt:r8>0.5</vt:r8></property>' +
			'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Due"><vt:filetime>2024-05-01T00:00:00Z</vt:filetime></property>' +
			'</Properties>';
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [{ name: 'S', xml: worksheet('') }],
				packageRels: [
					{
						id: 'rId2',
						type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
						target: 'docProps/core.xml',
					},
					{ id: 'rId3', type: 'extended-properties', target: 'docProps/app.xml' },
					{ id: 'rId4', type: 'custom-properties', target: 'docProps/custom.xml' },
				],
				parts: {
					'docProps/core.xml': core,
					'docProps/app.xml': app,
					'docProps/custom.xml': custom,
				},
			}),
		);
		expect(wb.properties).toEqual({
			title: 'Budget',
			subject: 'Finance',
			creator: 'Ann',
			keywords: 'q1, plan',
			description: 'Quarterly budget',
			lastModifiedBy: 'Ben',
			category: 'Reports',
			created: '2024-01-02T03:04:05Z',
			modified: '2024-02-03T04:05:06Z',
			application: 'Microsoft Excel',
			appVersion: '16.0300',
			company: 'ACME & Co',
			manager: 'Mia',
			custom: [
				{ name: 'Project', value: 'Apollo', type: 'lpwstr' },
				{ name: 'Version', value: 42, type: 'i4' },
				{ name: 'Approved', value: true, type: 'bool' },
				{ name: 'Ratio', value: 0.5, type: 'r8' },
				{ name: 'Due', value: '2024-05-01T00:00:00Z', type: 'filetime' },
			],
		});
	});

	it('records a warning for an unreadable properties part and continues', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [{ name: 'S', xml: worksheet('') }],
				packageRels: [{ id: 'rId2', type: 'extended-properties', target: 'docProps/app.xml' }],
				parts: { 'docProps/app.xml': '<Properties><Application>X' },
			}),
		);
		expect(wb.properties).toEqual({});
		expect(wb.warnings.map((w) => w.code)).toEqual(['unreadable-part']);
	});
});

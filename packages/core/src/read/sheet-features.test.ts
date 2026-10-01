import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { MAX_ROW_INDEX } from '../address/constants';
import type { Worksheet } from '../types/sheet';
import { readWorkbook } from './read-workbook';
import { buildWorkbook, worksheet, type RelSpec } from './test-utils/package-builder';

async function readSheet(
	before: string,
	after: string,
	rels?: RelSpec[],
	data = '',
): Promise<Worksheet> {
	const wb = await readWorkbook(
		await buildWorkbook({ sheets: [{ name: 'S', xml: worksheet(data, { before, after }), rels }] }),
	);
	return wb.sheets[0];
}

describe('worksheet features', () => {
	it('reads sheetPr, the sheet view with a frozen pane and the active-pane selection', async () => {
		const sheet = await readSheet(
			'<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0" summaryRight="false"/></sheetPr>' +
				'<dimension ref="A1:C3"/>' +
				'<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" showZeros="0" showFormulas="1" rightToLeft="1" ' +
				'zoomScale="85" tabSelected="1" view="pageLayout" topLeftCell="B2" defaultGridColor="0" colorId="10" workbookViewId="0">' +
				'<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>' +
				'<selection pane="topRight" activeCell="B1" sqref="B1"/>' +
				'<selection pane="bottomRight" activeCell="C5" sqref="C5:D6 F1"/>' +
				'</sheetView><sheetView workbookViewId="1"/></sheetViews>' +
				'<sheetFormatPr baseColWidth="10" defaultColWidth="12.5" defaultRowHeight="18" zeroHeight="1" outlineLevelRow="2" outlineLevelCol="1"/>',
			'',
		);
		expect(sheet.tabColor).toEqual({ rgb: 'FFFF0000' });
		expect(sheet.summaryBelow).toBe(false);
		expect(sheet.summaryRight).toBe(false);
		expect(sheet.view).toEqual({
			showGridLines: false,
			showRowColHeaders: false,
			showZeros: false,
			showFormulas: true,
			rightToLeft: true,
			zoomScale: 85,
			tabSelected: true,
			view: 'pageLayout',
			topLeftCell: { row: 1, col: 1 },
			gridColor: { indexed: 10 },
			pane: {
				xSplit: 1,
				ySplit: 2,
				topLeftCell: { row: 2, col: 1 },
				activePane: 'bottomRight',
				state: 'frozen',
			},
			selection: {
				activeCell: { row: 4, col: 2 },
				ranges: [
					{ top: 4, left: 2, bottom: 5, right: 3 },
					{ top: 0, left: 5, bottom: 0, right: 5 },
				],
				pane: 'bottomRight',
			},
		});
		expect(sheet.defaultRowHeight).toBe(18);
		expect(sheet.defaultColWidth).toBe(12.5);
		expect(sheet.baseColWidth).toBe(10);
		expect(sheet.zeroHeight).toBe(true);
		expect(sheet.outlineLevelRow).toBe(2);
		expect(sheet.outlineLevelCol).toBe(1);
	});

	it('uses defaults when there is no sheet view', async () => {
		const sheet = await readSheet('', '');
		expect(sheet.view).toEqual({
			showGridLines: true,
			showRowColHeaders: true,
			showZeros: true,
			showFormulas: false,
			rightToLeft: false,
			zoomScale: 100,
		});
		expect(sheet.defaultRowHeight).toBe(15);
		expect(sheet.merges).toEqual([]);
		expect(sheet.cols).toEqual([]);
	});

	it('reads columns (0-based), merges and protection', async () => {
		const sheet = await readSheet(
			'<cols><col min="3" max="5" width="20" customWidth="1" style="2" hidden="1"/><col min="1" max="1" width="9.140625" bestFit="1" outlineLevel="1" collapsed="1"/></cols>' +
				'<cols><col min="10" max="16384" width="5"/></cols>',
			'<sheetProtection algorithmName="SHA-512" hashValue="h" saltValue="s" spinCount="100000" sheet="1" objects="1" scenarios="1" formatCells="0" sort="0" selectLockedCells="1"/>' +
				'<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="D4:D9"/><mergeCell ref="Z1"/></mergeCells>',
		);
		expect(sheet.cols).toEqual([
			{ min: 0, max: 0, width: 9.140625, bestFit: true, outlineLevel: 1, collapsed: true },
			{ min: 2, max: 4, width: 20, customWidth: true, style: 2, hidden: true },
			{ min: 9, max: 16383, width: 5 },
		]);
		expect(sheet.merges).toEqual([
			{ top: 0, left: 0, bottom: 1, right: 1 },
			{ top: 3, left: 3, bottom: 8, right: 3 },
		]);
		expect(sheet.protection).toEqual({
			algorithmName: 'SHA-512',
			hashValue: 'h',
			saltValue: 's',
			spinCount: 100000,
			sheet: true,
			objects: true,
			scenarios: true,
			formatCells: false,
			selectLockedCells: true,
			sort: false,
		});
	});

	it('reads hyperlinks with external targets from the sheet relationships', async () => {
		const sheet = await readSheet(
			'',
			'<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Go" display="Site"/>' +
				'<hyperlink ref="B2:C3" location="\'Other Sheet\'!A1" display="Jump"/>' +
				'<hyperlink ref="D4" r:id="rId2"/></hyperlinks>',
			[
				{ id: 'rId1', type: 'hyperlink', target: 'https://example.com/?a=1&b=2', external: true },
				{ id: 'rId2', type: 'hyperlink', target: 'mailto:someone@example.com', external: true },
			],
		);
		expect(sheet.hyperlinks).toEqual([
			{
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'https://example.com/?a=1&b=2',
				display: 'Site',
				tooltip: 'Go',
			},
			{
				ref: { top: 1, left: 1, bottom: 2, right: 2 },
				location: "'Other Sheet'!A1",
				display: 'Jump',
			},
			{ ref: { top: 3, left: 3, bottom: 3, right: 3 }, target: 'mailto:someone@example.com' },
		]);
		expect(sheet.source?.rels).toHaveLength(2);
	});

	it('reads print options, margins, page setup and header / footer', async () => {
		const sheet = await readSheet(
			'',
			'<printOptions gridLines="1" headings="1" horizontalCentered="1"/>' +
				'<pageMargins left="0.25" right="0.25" top="1" bottom="1" header="0.5" footer="0.5"/>' +
				'<pageSetup paperSize="9" scale="80" fitToWidth="1" fitToHeight="0" firstPageNumber="3" useFirstPageNumber="1" ' +
				'orientation="landscape" pageOrder="overThenDown" blackAndWhite="1" draft="0" horizontalDpi="600" verticalDpi="600" r:id="rId9"/>' +
				'<headerFooter differentFirst="1"><oddHeader>&amp;C&amp;"Arial,Bold"Title</oddHeader><oddFooter>Page &amp;P of &amp;N</oddFooter><firstHeader>First</firstHeader></headerFooter>',
		);
		expect(sheet.printOptions).toEqual({
			gridLines: true,
			headings: true,
			horizontalCentered: true,
		});
		expect(sheet.pageMargins).toEqual({
			left: 0.25,
			right: 0.25,
			top: 1,
			bottom: 1,
			header: 0.5,
			footer: 0.5,
		});
		expect(sheet.pageSetup).toEqual({
			paperSize: 9,
			orientation: 'landscape',
			scale: 80,
			fitToWidth: 1,
			fitToHeight: 0,
			firstPageNumber: 3,
			useFirstPageNumber: true,
			pageOrder: 'overThenDown',
			blackAndWhite: true,
			draft: false,
			horizontalDpi: 600,
			verticalDpi: 600,
		});
		expect(sheet.headerFooter).toEqual({
			differentFirst: true,
			oddHeader: '&C&"Arial,Bold"Title',
			oddFooter: 'Page &P of &N',
			firstHeader: 'First',
		});
	});

	it('reads an autoFilter with every filter column form and a sort state', async () => {
		const sheet = await readSheet(
			'',
			'<autoFilter ref="A1:G20">' +
				'<filterColumn colId="0"><filters blank="1"><filter val="Apples"/><filter val="Pears"/>' +
				'<dateGroupItem year="2024" month="3" dateTimeGrouping="month"/></filters></filterColumn>' +
				'<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"/><customFilter val="*x*"/></customFilters></filterColumn>' +
				'<filterColumn colId="2"><top10 top="0" percent="1" val="5" filterVal="3.5"/></filterColumn>' +
				'<filterColumn colId="3"><dynamicFilter type="aboveAverage" val="12.5"/></filterColumn>' +
				'<filterColumn colId="4"><colorFilter dxfId="2" cellColor="0"/></filterColumn>' +
				'<filterColumn colId="5"><iconFilter iconSet="3Arrows" iconId="1"/></filterColumn>' +
				'<filterColumn colId="6" hiddenButton="1"/>' +
				'<sortState ref="A2:G20"><sortCondition descending="1" ref="B2:B20"/><sortCondition sortBy="cellColor" ref="C2:C20" dxfId="4"/></sortState>' +
				'</autoFilter>' +
				'<sortState ref="J2:K9" columnSort="1" caseSensitive="1"><sortCondition ref="J2:J9" customList="a,b"/></sortState>',
		);
		const filter = sheet.autoFilter!;
		expect(filter.ref).toEqual({ top: 0, left: 0, bottom: 19, right: 6 });
		expect(filter.columns).toEqual([
			{
				colId: 0,
				filters: {
					values: ['Apples', 'Pears'],
					blank: true,
					dateGroups: [{ dateTimeGrouping: 'month', year: 2024, month: 3 }],
				},
			},
			{
				colId: 1,
				customFilters: {
					and: true,
					filters: [{ operator: 'greaterThan', val: '10' }, { val: '*x*' }],
				},
			},
			{ colId: 2, top10: { top: false, percent: true, val: 5, filterVal: 3.5 } },
			{ colId: 3, dynamicFilter: { type: 'aboveAverage', val: 12.5 } },
			{ colId: 4, colorFilter: { dxfId: 2, cellColor: false } },
			{ colId: 5, iconFilter: { iconSet: '3Arrows', iconId: 1 } },
			{ colId: 6, hiddenButton: true },
		]);
		expect(filter.sortState).toEqual({
			ref: { top: 1, left: 0, bottom: 19, right: 6 },
			conditions: [
				{ ref: { top: 1, left: 1, bottom: 19, right: 1 }, descending: true },
				{ ref: { top: 1, left: 2, bottom: 19, right: 2 }, sortBy: 'cellColor', dxfId: 4 },
			],
		});
		expect(sheet.sortState).toEqual({
			ref: { top: 1, left: 9, bottom: 8, right: 10 },
			columnSort: true,
			caseSensitive: true,
			conditions: [{ ref: { top: 1, left: 9, bottom: 8, right: 9 }, customList: 'a,b' }],
		});
	});

	it('records feature warnings for pivot tables, OLE objects and controls', async () => {
		const bytes = await buildWorkbook({
			sheets: [
				{
					name: 'P',
					xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>', {
						after:
							'<mc:AlternateContent><mc:Choice Requires="x14"><controls><control shapeId="1025" r:id="rId3"/></controls></mc:Choice></mc:AlternateContent>' +
							'<oleObjects><oleObject progId="Word.Document.12" shapeId="1026" r:id="rId4"/></oleObjects>',
					}),
					rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }],
				},
			],
		});
		const wb = await readWorkbook(bytes);
		const codes = wb.warnings.map((w) => w.code).sort();
		expect(codes).toEqual(['form-control', 'ole-object', 'pivot-table']);
		expect(wb.warnings.every((w) => w.sheet === 'P')).toBe(true);
		expect(wb.sheets[0].cells.get(cellKey(0, 0))?.value).toBe(1);
	});

	it('parses whole-column hyperlinks and keeps unknown elements in the source root', async () => {
		const sheet = await readSheet(
			'',
			'<hyperlinks><hyperlink ref="A:A" location="Top"/></hyperlinks><ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/></ignoredErrors>',
		);
		expect(sheet.hyperlinks[0].ref).toEqual({ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: 0 });
		const names = sheet.source?.root?.children.map((c) =>
			typeof c === 'object' && 'name' in c ? c.name : '',
		);
		expect(names).toContain('ignoredErrors');
	});
});

import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { MAX_ROW_INDEX } from '../address/constants';
import { createWorksheet } from '../sheet/worksheet';
import type { Worksheet } from '../types/sheet';
import { parseXml } from '../xml';
import type { XmlElement } from '../xml/types';
import { CellMetadata } from './metadata';
import { SharedStringTable } from './shared-strings';
import { childNames, pick, pickAll } from './test-utils/package-check';
import { buildWorksheetXml, type WorksheetLinks, WORKSHEET_ORDER } from './worksheet-xml';

function build(sheet: Worksheet, root?: XmlElement, links: Partial<WorksheetLinks> = {}) {
	const res = buildWorksheetXml(
		sheet,
		root,
		{ hyperlinks: new Map(), ...links },
		{
			styleCount: 5,
			styles: { dxfs: () => [{}], internDxf: () => 0 },
			sheetData: { sst: SharedStringTable.empty(), styleCount: 5, metadata: new CellMetadata() },
			isNew: !root,
		},
	);
	return { ...res, root: parseXml(res.xml).root };
}

function richSheet(): Worksheet {
	const s = createWorksheet('Full', 1);
	s.cells.set(cellKey(0, 0), { value: 'Head' });
	s.cells.set(cellKey(4, 3), { value: 2 });
	s.tabColor = { rgb: 'FFFF0000' };
	s.summaryBelow = false;
	s.view = {
		...s.view,
		showGridLines: false,
		zoomScale: 150,
		tabSelected: true,
		pane: { xSplit: 1, ySplit: 2, state: 'frozen' },
		selection: {
			activeCell: { row: 3, col: 2 },
			ranges: [{ top: 3, left: 2, bottom: 4, right: 3 }],
		},
	};
	s.cols = [
		{ min: 0, max: 0, width: 20, customWidth: true },
		{ min: 1, max: 1, width: 12, customWidth: true },
		{ min: 2, max: 3, width: 12, customWidth: true },
		{ min: 5, max: 5, hidden: true, outlineLevel: 1 },
	];
	s.merges = [
		{ top: 0, left: 0, bottom: 0, right: 2 },
		{ top: 0, left: 1, bottom: 1, right: 1 },
		{ top: 7, left: 7, bottom: 7, right: 7 },
	];
	s.protection = { sheet: true, formatCells: false, algorithmName: 'SHA-512', spinCount: 100000 };
	s.autoFilter = {
		ref: { top: 0, left: 0, bottom: 4, right: 3 },
		columns: [{ colId: 1, filters: { values: ['a', 'b'] } }],
	};
	s.conditionalFormats = [
		{
			ranges: [{ top: 1, left: 0, bottom: 4, right: 0 }],
			rules: [{ type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['5'], dxfId: 0 }],
		},
	];
	s.dataValidations = [
		{
			ranges: [{ top: 1, left: 1, bottom: 9, right: 1 }],
			type: 'list',
			formula1: '"Yes,No"',
			allowBlank: true,
			inCellDropdown: false,
		},
	];
	s.hyperlinks = [
		{ ref: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com', tooltip: 'Go' },
		{
			ref: { top: 1, left: 0, bottom: 1, right: 0 },
			location: "'Other Sheet'!A1",
			display: 'Other',
		},
		{ ref: { top: 2, left: 0, bottom: 2, right: 0 }, target: '#Data!B2' },
		{ ref: { top: 3, left: 0, bottom: 3, right: 0 } },
	];
	s.printOptions = { gridLines: true, horizontalCentered: true };
	s.pageMargins = { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.4, footer: 0.4 };
	s.pageSetup = { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0 };
	s.headerFooter = {
		oddHeader: '&CReport & Summary',
		oddFooter: '&P of &N',
		differentFirst: true,
		firstHeader: 'First',
	};
	s.sparklineGroups = [
		{
			type: 'column',
			sparklines: [{ formula: 'Full!A2:D2', location: { row: 1, col: 4 } }],
			markers: true,
		},
	];
	return s;
}

describe('worksheet assembly', () => {
	it('writes every model element in CT_Worksheet order', () => {
		const s = richSheet();
		const link = s.hyperlinks[0];
		const { root, droppedMerges } = build(s, undefined, {
			hyperlinks: new Map([[link, 'rId3']]),
			drawing: 'rId1',
			legacyDrawing: 'rId2',
			tableParts: ['rId4', 'rId5'],
		});
		const names = childNames(root);
		expect(names).toEqual([
			'sheetPr',
			'dimension',
			'sheetViews',
			'sheetFormatPr',
			'cols',
			'sheetData',
			'sheetProtection',
			'autoFilter',
			'mergeCells',
			'conditionalFormatting',
			'dataValidations',
			'hyperlinks',
			'printOptions',
			'pageMargins',
			'pageSetup',
			'headerFooter',
			'drawing',
			'legacyDrawing',
			'tableParts',
			'extLst',
		]);
		const ranks = names.map((n) => WORKSHEET_ORDER.indexOf(n as (typeof WORKSHEET_ORDER)[number]));
		expect([...ranks].sort((a, b) => a - b)).toEqual(ranks);
		expect(droppedMerges).toBe(2);
		expect(pick(root, 'sheetPr', 'tabColor')?.attrs.rgb).toBe('FFFF0000');
		expect(pick(root, 'sheetPr', 'outlinePr')?.attrs).toEqual({ summaryBelow: '0' });
		expect(pick(root, 'sheetPr', 'pageSetUpPr')?.attrs).toEqual({ fitToPage: '1' });
		expect(pick(root, 'dimension')?.attrs.ref).toBe('A1:D5');
		const view = pick(root, 'sheetViews', 'sheetView')!;
		expect(view.attrs).toMatchObject({
			showGridLines: '0',
			tabSelected: '1',
			zoomScale: '150',
			workbookViewId: '0',
		});
		expect(pick(view, 'pane')?.attrs).toEqual({
			xSplit: '1',
			ySplit: '2',
			topLeftCell: 'B3',
			activePane: 'bottomRight',
			state: 'frozen',
		});
		expect(pick(view, 'selection')?.attrs).toEqual({
			pane: 'bottomRight',
			activeCell: 'C4',
			sqref: 'C4:D5',
		});
		expect(pickAll(pick(root, 'cols'), 'col').map((c) => c.attrs)).toEqual([
			{ min: '1', max: '1', width: '20', customWidth: '1' },
			{ min: '2', max: '4', width: '12', customWidth: '1' },
			{ min: '6', max: '6', width: '9.140625', hidden: '1', outlineLevel: '1' },
		]);
		expect(pick(root, 'sheetFormatPr')?.attrs).toEqual({
			defaultRowHeight: '15',
			outlineLevelCol: '1',
		});
		expect(pick(root, 'sheetProtection')?.attrs).toMatchObject({
			sheet: '1',
			formatCells: '0',
			algorithmName: 'SHA-512',
		});
		expect(pick(root, 'autoFilter', 'filterColumn', 'filters')).toBeDefined();
		expect(pick(root, 'mergeCells')?.attrs.count).toBe('1');
		expect(pick(root, 'dataValidations', 'dataValidation')?.attrs).toMatchObject({
			type: 'list',
			showDropDown: '1',
			sqref: 'B2:B10',
		});
		const links = pickAll(pick(root, 'hyperlinks'), 'hyperlink');
		expect(links.map((h) => h.attrs)).toEqual([
			{ ref: 'A1', 'r:id': 'rId3', tooltip: 'Go' },
			{ ref: 'A2', location: "'Other Sheet'!A1", display: 'Other' },
			{ ref: 'A3', location: 'Data!B2' },
		]);
		expect(pick(root, 'pageSetup')?.attrs).toEqual({
			paperSize: '9',
			fitToWidth: '1',
			fitToHeight: '0',
			orientation: 'landscape',
		});
		const hf = pick(root, 'headerFooter')!;
		expect(hf.attrs).toEqual({ differentFirst: '1' });
		expect(childNames(hf)).toEqual(['oddHeader', 'oddFooter', 'firstHeader']);
		expect(pick(hf, 'oddHeader')?.children[0]).toBe('&CReport & Summary');
		expect(pickAll(pick(root, 'tableParts'), 'tablePart').map((t) => t.attrs['r:id'])).toEqual([
			'rId4',
			'rId5',
		]);
		expect(pick(root, 'extLst', 'ext')?.attrs.uri).toBe('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}');
	});

	it('reuses the original root: unknown elements, attributes and extensions survive', () => {
		const original = parseXml(
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{1}">' +
				'<sheetPr codeName="Sheet1"><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:Z99"/>' +
				'<sheetViews><sheetView zoomScaleNormal="120" zoomScale="120" workbookViewId="0"><selection activeCell="Z9" sqref="Z9"/></sheetView><sheetView workbookViewId="1"/></sheetViews>' +
				'<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData/>' +
				'<protectedRanges><protectedRange name="p" sqref="A1"/></protectedRanges>' +
				'<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>' +
				'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
				'<pageSetup paperSize="9" copies="2" r:id="rId1"/><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>' +
				'<legacyDrawingHF r:id="rId9"/>' +
				'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls><control shapeId="1025" r:id="rId7"/></controls></mc:Choice></mc:AlternateContent>' +
				'<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList/></ext></extLst>' +
				'</worksheet>',
		).root;
		const s = createWorksheet('S', 1);
		s.cells.set(0, { value: 1 });
		s.view = { ...s.view, zoomScale: 90 };
		s.pageSetup = { orientation: 'portrait' };
		s.dataValidations = [
			{
				ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
				type: 'list',
				formula1: 'Lists!$A$1:$A$3',
				fromExtension: true,
			},
		];
		const { root } = build(s, original, { drawing: 'rId20', tableParts: [] });
		expect(root.attrs['xr:uid']).toBe('{1}');
		expect(childNames(root)).toEqual([
			'sheetPr',
			'dimension',
			'sheetViews',
			'sheetFormatPr',
			'sheetData',
			'protectedRanges',
			'pageMargins',
			'pageSetup',
			'rowBreaks',
			'drawing',
			'legacyDrawingHF',
			'AlternateContent',
			'extLst',
		]);
		expect(pick(root, 'sheetPr')?.attrs.codeName).toBe('Sheet1');
		expect(pick(root, 'sheetPr', 'pageSetUpPr')).toBeDefined();
		expect(pick(root, 'dimension')?.attrs.ref).toBe('A1');
		const views = pickAll(pick(root, 'sheetViews'), 'sheetView');
		expect(views).toHaveLength(2);
		expect(views[0].attrs).toEqual({ zoomScaleNormal: '90', zoomScale: '90', workbookViewId: '0' });
		expect(pick(root, 'sheetFormatPr')?.attrs['x14ac:dyDescent']).toBe('0.25');
		expect(pick(root, 'pageSetup')?.attrs).toEqual({
			copies: '2',
			'r:id': 'rId1',
			orientation: 'portrait',
		});
		const exts = pickAll(pick(root, 'extLst'), 'ext').map((e) => e.attrs.uri);
		expect(exts).toEqual([
			'{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}',
			'{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
		]);
	});

	it('inserts new elements before mc:AlternateContent blocks that wrap later elements', () => {
		const original = parseXml(
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
				'<sheetData/><mc:AlternateContent><mc:Choice Requires="x14"><controls/></mc:Choice></mc:AlternateContent><tableParts count="0"/></worksheet>',
		).root;
		const s = createWorksheet('S', 1);
		s.headerFooter = { oddFooter: 'x' };
		const { root } = build(s, original, { drawing: 'rId1', legacyDrawing: 'rId2', tableParts: [] });
		const names = childNames(root);
		expect(names.indexOf('drawing')).toBeLessThan(names.indexOf('AlternateContent'));
		expect(names.indexOf('legacyDrawing')).toBeLessThan(names.indexOf('AlternateContent'));
		expect(names.indexOf('headerFooter')).toBeLessThan(names.indexOf('drawing'));
		expect(names).not.toContain('tableParts');
	});

	it('normalises a prefixed source root', () => {
		const original = parseXml(
			'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/><x:pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/></x:worksheet>',
		).root;
		const s = createWorksheet('S', 1);
		s.cells.set(0, { value: 'x' });
		const { root, xml } = build(s, original);
		expect(root.name).toBe('worksheet');
		expect(root.attrs.xmlns).toBe('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		expect(childNames(root).filter((n) => n === 'sheetData')).toHaveLength(1);
		expect(xml).toContain('<x:pageMargins');
	});

	it('writes whole-column / whole-row ranges in A1 forms Excel accepts', () => {
		const s = createWorksheet('S', 1);
		s.conditionalFormats = [
			{
				ranges: [{ top: 0, left: 1, bottom: MAX_ROW_INDEX, right: 1 }],
				rules: [{ type: 'expression', priority: 1, formulas: ['B1>0'] }],
			},
		];
		s.merges = [{ top: 0, left: 0, bottom: 1, right: 0 }];
		const { root } = build(s);
		expect(pick(root, 'conditionalFormatting')?.attrs.sqref).toBe('B:B');
		expect(pick(root, 'mergeCells', 'mergeCell')?.attrs.ref).toBe('A1:A2');
	});
});

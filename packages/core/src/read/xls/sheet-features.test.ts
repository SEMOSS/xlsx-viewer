import { describe, expect, it } from 'vitest';

import { BOF_TYPE, RT } from './biff-opcodes';
import { readXlsWorkbook } from './index';
import { buildXls, type SheetSpec, window1, xf } from './test-utils/biff-builder';
import { record, u16, u32, xlString } from './test-utils/bytes';
import {
	calcMode,
	colInfo,
	defaultRowHeight,
	defColWidth,
	header,
	margin,
	mergedCells,
	number,
	pane,
	protect,
	row,
	scl,
	selection,
	window2,
	wsbool,
} from './test-utils/sheet-builder';

function readSheets(sheets: SheetSpec[], globals: number[][] = []) {
	return readXlsWorkbook(buildXls(globals, sheets));
}

function readSheet(records: number[][], globals: number[][] = []) {
	return readSheets([{ name: 'Sheet1', records }], globals).sheets[0];
}

describe('rows and columns', () => {
	it('reads row heights, hidden rows, outline levels and row formats', () => {
		const sheet = readSheet(
			[
				defaultRowHeight(255),
				row(0, { height: 255 }),
				row(1, { height: 400, custom: true }),
				row(2, { height: 300 }),
				row(3, { hidden: true, custom: true }),
				row(4, { level: 2, collapsed: true }),
				row(5, { xf: 16 }),
			],
			[xf({ numFmt: 10 })],
		);
		expect(sheet.defaultRowHeight).toBe(12.75);
		expect(sheet.rows.get(0)).toBeUndefined();
		expect(sheet.rows.get(1)).toEqual({ height: 20, customHeight: true });
		expect(sheet.rows.get(2)).toEqual({ height: 15 });
		expect(sheet.rows.get(3)).toEqual({ height: 12.75, customHeight: true, hidden: true });
		expect(sheet.rows.get(4)).toEqual({ outlineLevel: 2, collapsed: true });
		expect(sheet.rows.get(5)).toEqual({ style: 1, customFormat: true });
		expect(sheet.outlineLevelRow).toBe(2);
	});

	it('reads column widths, hidden columns, styles and the default width', () => {
		const sheet = readSheet(
			[
				defColWidth(10),
				colInfo(0, 0, 2340),
				colInfo(1, 3, 5120, { custom: true, xf: 16 }),
				colInfo(5, 5, 0, { hidden: true, level: 1 }),
				colInfo(6, 300, 1000),
			],
			[xf({ fill: { fls: 1, fore: 10 } })],
		);
		expect(sheet.baseColWidth).toBe(10);
		expect(sheet.cols).toEqual([
			{ min: 0, max: 0, width: 9.140625 },
			{ min: 1, max: 3, width: 20, customWidth: true, style: 1 },
			{ min: 5, max: 5, width: 0, hidden: true, outlineLevel: 1 },
			{ min: 6, max: 300, width: 3.90625 },
		]);
		expect(sheet.outlineLevelCol).toBe(1);
	});

	it('reads hidden-by-default rows and outline summary positions', () => {
		const sheet = readSheet([defaultRowHeight(300, 0x02), wsbool(0x0401)]);
		expect(sheet.zeroHeight).toBe(true);
		expect(sheet.summaryBelow).toBe(false);
		expect(sheet.summaryRight).toBe(false);
	});

	it('reads merged cells and drops single-cell merges', () => {
		const sheet = readSheet([
			mergedCells([
				[0, 1, 0, 2],
				[3, 3, 3, 3],
			]),
			mergedCells([[5, 9, 1, 1]]),
		]);
		expect(sheet.merges).toEqual([
			{ top: 0, left: 0, bottom: 1, right: 2 },
			{ top: 5, left: 1, bottom: 9, right: 1 },
		]);
	});
});

describe('sheet view', () => {
	it('reads gridlines, headers, zeros, zoom and the scroll position', () => {
		const sheet = readSheet([
			window2({ flags: 0x0200 | 0x0040, top: 4, left: 2, grid: 10 }),
			scl(3, 2),
		]);
		expect(sheet.view).toEqual({
			showGridLines: false,
			showRowColHeaders: false,
			showZeros: false,
			showFormulas: false,
			rightToLeft: true,
			zoomScale: 150,
			tabSelected: true,
			topLeftCell: { row: 4, col: 2 },
			gridColor: { indexed: 10 },
		});
	});

	it('reads frozen panes with the selection of the active pane', () => {
		const sheet = readSheet([
			window2({ flags: 0x06be | 0x0100 }),
			pane(1, 2, 2, 1, 0),
			selection(3, 0, 0, [[0, 0, 0, 0]]),
			selection(0, 5, 3, [[5, 6, 3, 4]]),
		]);
		expect(sheet.view.pane).toEqual({
			xSplit: 1,
			ySplit: 2,
			topLeftCell: { row: 2, col: 1 },
			activePane: 'bottomRight',
			state: 'frozen',
		});
		expect(sheet.view.selection).toEqual({
			activeCell: { row: 5, col: 3 },
			ranges: [{ top: 5, left: 3, bottom: 6, right: 4 }],
			pane: 'bottomRight',
		});
	});

	it('reads split panes and page break preview zoom', () => {
		const sheet = readSheet([window2({ flags: 0x06b6 | 0x0800, zoom: 0 }), pane(2400, 0, 0, 3, 1)]);
		expect(sheet.view.view).toBe('pageBreakPreview');
		expect(sheet.view.pane).toEqual({
			xSplit: 2400,
			ySplit: 0,
			topLeftCell: { row: 0, col: 3 },
			activePane: 'topRight',
			state: 'split',
		});
	});

	it('reads the tab color from SHEETEXT', () => {
		const sheetExt = record(RT.SHEETEXT, [
			...u16(RT.SHEETEXT),
			...u16(0),
			...new Array<number>(8).fill(0),
			...u32(20),
			...u32(0x0a),
		]);
		expect(readSheet([sheetExt]).tabColor).toEqual({ indexed: 10 });
	});
});

describe('sheets', () => {
	it('reads sheet states, skips chart and macro sheets and remaps the active tab', () => {
		const wb = readSheets(
			[
				{ name: 'Visible', records: [number(0, 0, 15, 1)] },
				{ name: 'Chart1', records: [], type: 2, bofType: BOF_TYPE.chart },
				{ name: 'Hidden', records: [], state: 'hidden' },
				{ name: 'Macro1', records: [], type: 1, bofType: BOF_TYPE.macro },
				{ name: 'VeryHidden', records: [], state: 'veryHidden' },
				{ name: 'Dialog', records: [wsbool(0x04d1)] },
			],
			[window1({ activeTab: 2, firstTab: 2, flags: 0x28, tabRatio: 750 })],
		);
		expect(wb.sheets.map((s) => [s.name, s.state, s.sheetId])).toEqual([
			['Visible', 'visible', 1],
			['Hidden', 'hidden', 2],
			['VeryHidden', 'veryHidden', 3],
		]);
		expect(wb.view).toEqual({
			activeTab: 1,
			firstSheet: 1,
			showVerticalScroll: false,
			tabRatio: 750,
		});
		const skipped = wb.warnings.filter((w) => w.code === 'sheet-skipped').map((w) => w.sheet);
		expect(skipped).toEqual(['Chart1', 'Macro1', 'Dialog']);
		expect(wb.dirty).toBe(true);
		expect(wb.sheets.every((s) => s.dirty === true && s.source === undefined)).toBe(true);
		expect(wb.sourcePackage).toBeUndefined();
	});

	it('adds an empty sheet when nothing can be imported', () => {
		const wb = readSheets([{ name: 'Chart1', records: [], type: 2, bofType: BOF_TYPE.chart }]);
		expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1']);
		expect(wb.warnings.map((w) => w.code)).toContain('no-worksheets');
	});

	it('reports embedded charts, pictures and shapes', () => {
		const obj = (type: number, id: number): number[] =>
			record(RT.OBJ, [
				...u16(0x15),
				...u16(0x12),
				...u16(type),
				...u16(id),
				...u16(0),
				...new Array<number>(12).fill(0),
			]);
		const wb = readSheets([
			{ name: 'S', records: [obj(5, 1), obj(8, 2), obj(2, 3), obj(0x14, 4)] },
		]);
		expect(wb.warnings.map((w) => w.code).sort()).toEqual([
			'chart-skipped',
			'drawing-skipped',
			'image-skipped',
		]);
	});
});

describe('print settings, protection and calculation', () => {
	it('reads margins, page setup, headers and print options', () => {
		const setup = record(RT.SETUP, [
			...u16(9),
			...u16(80),
			...u16(3),
			...u16(1),
			...u16(0),
			...u16(0x80 | 0x01),
			...u16(600),
			...u16(600),
			...new Array<number>(16).fill(0),
			...u16(1),
		]);
		const sheet = readSheet([
			margin(RT.LEFTMARGIN, 0.5),
			margin(RT.TOPMARGIN, 1.25),
			setup,
			header('&CPage &P'),
			record(RT.FOOTER, []),
			record(RT.PRINTGRIDLINES, u16(1)),
			record(RT.HCENTER, u16(1)),
			record(RT.VCENTER, u16(0)),
		]);
		expect(sheet.pageMargins).toEqual({
			left: 0.5,
			right: 0.75,
			top: 1.25,
			bottom: 1,
			header: 0,
			footer: 0,
		});
		expect(sheet.pageSetup).toEqual({
			paperSize: 9,
			scale: 80,
			orientation: 'landscape',
			horizontalDpi: 600,
			verticalDpi: 600,
			fitToHeight: 0,
			firstPageNumber: 3,
			useFirstPageNumber: true,
			pageOrder: 'overThenDown',
		});
		expect(sheet.headerFooter).toEqual({ oddHeader: '&CPage &P' });
		expect(sheet.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
	});

	it('reads sheet and workbook protection and calculation settings', () => {
		const wb = readSheets(
			[
				{
					name: 'S',
					records: [
						protect(),
						record(RT.OBJPROTECT, u16(1)),
						record(RT.PASSWORD, u16(0xcc3d)),
						calcMode(0),
						record(RT.ITERATION, u16(1)),
						record(RT.CALCCOUNT, u16(50)),
					],
				},
			],
			[
				record(RT.PROTECT, u16(1)),
				record(RT.PASSWORD, u16(0x83af)),
				record(RT.WRITEACCESS, [...xlString('Jane Doe'), ...new Array<number>(20).fill(0x20)]),
			],
		);
		expect(wb.sheets[0].protection).toEqual({ sheet: true, objects: true, password: 'CC3D' });
		expect(wb.protection).toEqual({ lockStructure: true, workbookPassword: '83AF' });
		expect(wb.calc).toEqual({ calcMode: 'manual', iterate: true, iterateCount: 50 });
		expect(wb.properties.lastModifiedBy).toBe('Jane Doe');
	});
});

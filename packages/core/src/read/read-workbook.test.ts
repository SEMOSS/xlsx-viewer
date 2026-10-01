import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { OpcPackage } from '../package/opc-package';
import type { Workbook } from '../types/workbook';
import { XlsxFormatError, XlsxPasswordError, readWorkbook } from './index';
import {
	buildWorkbook,
	relsXml,
	worksheet,
	zipParts,
	workbookParts,
} from './test-utils/package-builder';

const FIXTURES = join(__dirname, '../__tests__/fixtures/real');

function fixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(FIXTURES, name)));
}

function value(wb: Workbook, sheet: number, row: number, col: number): unknown {
	return wb.sheets[sheet].cells.get(cellKey(row, col))?.value;
}

describe('readWorkbook: real Excel files', () => {
	it('reads every fixture without warnings about unreadable parts', async () => {
		const files = readdirSync(FIXTURES).filter((f) => f.endsWith('.xlsx'));
		expect(files.length).toBeGreaterThan(5);
		for (const f of files) {
			const wb = await readWorkbook(fixture(f));
			expect(wb.sheets.length, f).toBeGreaterThan(0);
			expect(
				wb.warnings.filter((w) => w.code === 'unreadable-part'),
				f,
			).toEqual([]);
			expect(wb.sourcePackage).toBeInstanceOf(OpcPackage);
			expect(wb.dirty).toBe(false);
			for (const s of wb.sheets) {
				expect(s.dirty).toBe(false);
				expect(s.drawingsDirty).toBe(false);
				expect(s.commentsDirty).toBe(false);
				expect(s.source?.partPath).toBe('xl/worksheets/sheet1.xml');
			}
		}
	});

	it('reads the HR workbook (CJK strings with phonetic properties, table, styles)', async () => {
		const wb = await readWorkbook(fixture('issue-132-hr-deck--Workbook1.xlsx'));
		const sheet = wb.sheets[0];
		expect(sheet.name).toBe('Sheet1');
		expect(sheet.sheetId).toBe(1);
		expect(sheet.kind).toBe('worksheet');
		expect(sheet.state).toBe('visible');
		expect(value(wb, 0, 0, 0)).toBe(' ');
		expect(value(wb, 0, 0, 1)).toBe('列1');
		expect(value(wb, 0, 1, 0)).toBe('硕士');
		expect(value(wb, 0, 5, 0)).toBe('高中');
		expect(sheet.cells.get(cellKey(1, 1))).toEqual({ value: 0.04, style: 1 });
		expect(value(wb, 0, 5, 1)).toBeCloseTo(0.07, 12);
		expect(sheet.cells.size).toBe(12);
		expect(sheet.defaultRowHeight).toBe(13.5);
		expect(sheet.view.zoomScale).toBe(160);
		expect(sheet.view.tabSelected).toBe(true);
		expect(sheet.view.selection?.activeCell).toEqual({ row: 6, col: 1 });
		expect(sheet.pageMargins).toEqual({
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		});
		expect(sheet.pageSetup).toEqual({ paperSize: 9, orientation: 'portrait' });

		expect(sheet.tables).toHaveLength(1);
		const table = sheet.tables[0];
		expect(table.name).toBe('表1');
		expect(table.ref).toEqual({ top: 0, left: 0, bottom: 5, right: 1 });
		expect(table.columns.map((c) => c.name)).toEqual([' ', '列1']);
		expect(table.style).toEqual({
			showFirstColumn: false,
			showLastColumn: false,
			showRowStripes: true,
			showColumnStripes: false,
		});
		expect(table.partPath).toBe('xl/tables/table1.xml');
		expect(table.dirty).toBe(false);
		expect(table.raw?.name).toBe('table');

		const styles = wb.styles;
		expect(styles.fonts[0]).toEqual({
			name: '宋体',
			size: 11,
			color: { theme: 1 },
			family: 2,
			charset: 134,
			scheme: 'minor',
		});
		expect(styles.cellXfs[1]).toEqual({
			numFmtId: 9,
			fontId: 0,
			fillId: 0,
			borderId: 0,
			xfId: 0,
			alignment: { vertical: 'center' },
			applyNumberFormat: true,
		});
		expect(styles.cellStyles[0]).toEqual({ name: '常规', xfId: 0, builtinId: 0 });
		expect(styles.raw?.cellXfs).toHaveLength(2);
		expect(styles.raw?.root?.name).toBe('styleSheet');
		expect(styles.defaultTableStyle).toBe('TableStyleMedium2');

		expect(wb.properties.creator).toBe('kingpub');
		expect(wb.properties.created).toBe('2017-12-16T13:48:55Z');
		expect(wb.properties.application).toBe('Microsoft Excel');
		expect(wb.properties.appVersion).toBe('15.0300');
		expect(wb.properties.company).toBeUndefined();
		expect(wb.calc.calcId).toBe(0);
		expect(wb.hasMacros).toBe(false);
	});

	it('reads columns, numbers and the Aptos theme of a modern workbook', async () => {
		const wb = await readWorkbook(fixture('chart-label-callouts--Microsoft_Excel_Worksheet.xlsx'));
		const sheet = wb.sheets[0];
		expect(sheet.cols).toEqual([
			{ min: 0, max: 0, width: 11.28515625, customWidth: true },
			{ min: 1, max: 3, width: 10.28515625, customWidth: true },
		]);
		expect(value(wb, 0, 0, 1)).toBe('Series 1');
		expect(value(wb, 0, 2, 2)).toBe(4.4);
		expect(value(wb, 0, 4, 3)).toBe(5);
		expect(sheet.tables[0].displayName).toBe('Table1');
		expect(wb.theme.name).toBe('Office Theme');
		expect(wb.theme.fonts.minorLatin).toBe('Aptos Narrow');
		expect(wb.theme.fonts.majorLatin).toBe('Aptos Display');
		expect(wb.theme.colors.accent1).toBe('156082');
		expect(wb.theme.colors.dk1).toBe('000000');
		expect(wb.theme.rawXml).toContain('<a:clrScheme name="Office">');
		expect(wb.calc.calcId).toBe(191029);
	});

	it('keeps the worksheet root without sheetData children for round-trip', async () => {
		const wb = await readWorkbook(fixture('ole-editable--Microsoft_Excel_Worksheet.xlsx'));
		const root = wb.sheets[0].source?.root;
		expect(root?.name).toBe('worksheet');
		const names = root?.children.map((c) => (typeof c === 'object' && 'name' in c ? c.name : ''));
		expect(names).toEqual(['dimension', 'sheetViews', 'sheetFormatPr', 'sheetData', 'pageMargins']);
		const sheetData = root?.children[3];
		expect(
			sheetData && typeof sheetData === 'object' && 'children' in sheetData && sheetData.children,
		).toEqual([]);
		expect(value(wb, 0, 0, 0)).toBe('Revenue');
		expect(value(wb, 0, 0, 1)).toBe(42);
	});
});

describe('readWorkbook: container errors', () => {
	function cfb(withEncryption: boolean): Uint8Array {
		const bytes = new Uint8Array(1024);
		bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
		if (withEncryption) {
			const name = 'EncryptionInfo';
			for (let i = 0; i < name.length; i++) {
				bytes[512 + i * 2] = name.charCodeAt(i);
			}
		}
		return bytes;
	}

	it('rejects encrypted packages with a password error', async () => {
		const err = await readWorkbook(cfb(true)).catch((e: unknown) => e);
		expect(err).toBeInstanceOf(XlsxPasswordError);
		expect((err as XlsxPasswordError).reason).toBe('unsupported');
	});

	it('reads compound files as legacy .xls and rejects broken ones as format errors', async () => {
		await expect(readWorkbook(cfb(false))).rejects.toThrow(XlsxFormatError);
		await expect(readWorkbook(cfb(false))).rejects.toThrow(/\.xls file is damaged/);
	});

	it('rejects non-ZIP input and ZIPs without a workbook', async () => {
		await expect(readWorkbook(new TextEncoder().encode('a,b,c\n1,2,3'))).rejects.toThrow(
			XlsxFormatError,
		);
		const notExcel = await zipParts({ 'hello.txt': 'hi' });
		await expect(readWorkbook(notExcel)).rejects.toThrow(/does not contain a workbook/);
		const truncated = (await buildWorkbook({ sheets: [] })).slice(0, 40);
		await expect(readWorkbook(truncated)).rejects.toThrow(XlsxFormatError);
	});

	it('names unsupported package formats', async () => {
		const ods = await zipParts({ mimetype: 'application/vnd.oasis.opendocument.spreadsheet' });
		await expect(readWorkbook(ods)).rejects.toThrow(/OpenDocument/);
		const parts = workbookParts({ sheets: [] });
		parts['_rels/.rels'] = relsXml([
			{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.bin' },
		]);
		parts['xl/workbook.bin'] = new Uint8Array([0x83, 0x01, 0x00]);
		await expect(readWorkbook(await zipParts(parts))).rejects.toThrow(/\.xlsb/);
		const docx = await zipParts({
			'_rels/.rels': relsXml([{ id: 'rId1', type: 'officeDocument', target: 'word/document.xml' }]),
			'word/document.xml': '<w:document xmlns:w="urn:w"/>',
		});
		await expect(readWorkbook(docx)).rejects.toThrow(/<w:document>/);
	});

	it('accepts an ArrayBuffer', async () => {
		const bytes = await buildWorkbook({ sheets: [{ name: 'A', xml: worksheet('') }] });
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		const wb = await readWorkbook(buffer);
		expect(wb.sheets[0].name).toBe('A');
	});
});

describe('readWorkbook: workbook part', () => {
	it('reads sheets, kinds, states, defined names, views, calc and protection', async () => {
		const parts = workbookParts({
			sheets: [
				{ name: 'Data', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>') },
				{ name: 'Hidden', xml: worksheet(''), state: 'hidden' },
				{ name: 'Secret', xml: worksheet(''), state: 'veryHidden' },
			],
			workbookBeforeSheets:
				'<workbookPr date1904="1"/><workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="abc" workbookSaltValue="salt" workbookSpinCount="100000"/>' +
				'<bookViews><workbookView activeTab="1" firstSheet="0" showSheetTabs="0" tabRatio="600"/></bookViews>',
			workbookAfterSheets:
				'<definedNames>' +
				'<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>' +
				'<definedName name="Rate" comment="Tax rate">0.2</definedName>' +
				'<definedName name="NewFn" hidden="1">_xlfn.XLOOKUP(1,Data!A:A,Data!B:B)</definedName>' +
				'</definedNames>' +
				'<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.01"/>',
		});
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets.map((s) => [s.name, s.state, s.sheetId])).toEqual([
			['Data', 'visible', 1],
			['Hidden', 'hidden', 2],
			['Secret', 'veryHidden', 3],
		]);
		expect(wb.date1904).toBe(true);
		expect(wb.view).toEqual({ activeTab: 1, firstSheet: 0, showSheetTabs: false, tabRatio: 600 });
		expect(wb.calc).toEqual({
			calcId: 191029,
			calcMode: 'manual',
			fullCalcOnLoad: true,
			iterate: true,
			iterateCount: 50,
			iterateDelta: 0.01,
		});
		expect(wb.protection).toEqual({
			lockStructure: true,
			workbookAlgorithmName: 'SHA-512',
			workbookHashValue: 'abc',
			workbookSaltValue: 'salt',
			workbookSpinCount: 100000,
		});
		expect(wb.definedNames).toEqual([
			{ name: '_xlnm.Print_Area', formula: 'Data!$A$1:$B$2', localSheet: 0, builtin: true },
			{ name: 'Rate', formula: '0.2', comment: 'Tax rate' },
			{ name: 'NewFn', formula: 'XLOOKUP(1,Data!A:A,Data!B:B)', hidden: true },
		]);
		// Without a styles or theme part the defaults are used.
		expect(wb.styles.cellXfs).toHaveLength(1);
		expect(wb.theme.colors.accent1).toBe('4472C4');
	});

	it('records a warning and an empty sheet for a missing or broken sheet part', async () => {
		const parts = workbookParts({
			sheets: [
				{ name: 'Good', xml: worksheet('<row r="1"><c r="A1"><v>7</v></c></row>') },
				{ name: 'Broken', xml: '<worksheet><sheetData><row></sheetData>' },
			],
		});
		const wb = await readWorkbook(await zipParts(parts));
		expect(value(wb, 0, 0, 0)).toBe(7);
		expect(wb.sheets[1].name).toBe('Broken');
		expect(wb.sheets[1].cells.size).toBe(0);
		expect(wb.warnings.some((w) => w.code === 'unreadable-part' && w.sheet === 'Broken')).toBe(
			true,
		);
	});

	it('detects macros and workbook-level features', async () => {
		const parts = workbookParts({
			sheets: [{ name: 'S', xml: worksheet('') }],
			workbookAfterSheets:
				'<externalReferences><externalReference r:id="rIdX"/></externalReferences>',
			workbookRels: [
				{
					id: 'rIdVba',
					type: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
					target: 'vbaProject.bin',
				},
				{ id: 'rIdX', type: 'externalLink', target: 'externalLinks/externalLink1.xml' },
				{ id: 'rIdC', type: 'connections', target: 'connections.xml' },
			],
			parts: { 'xl/vbaProject.bin': new Uint8Array([1, 2, 3]) },
		});
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.hasMacros).toBe(true);
		const codes = wb.warnings.map((w) => w.code);
		expect(codes).toContain('external-link');
		expect(codes).toContain('external-data');
	});

	it('finds the workbook through content types when the package relationship is missing', async () => {
		const parts = workbookParts({ sheets: [{ name: 'Only', xml: worksheet('') }] });
		parts['_rels/.rels'] = relsXml([]);
		const wb = await readWorkbook(await zipParts(parts));
		expect(wb.sheets[0].name).toBe('Only');
	});
});

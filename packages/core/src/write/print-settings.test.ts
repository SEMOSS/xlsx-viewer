import { describe, expect, it } from 'vitest';

import { readWorkbook } from '../read/read-workbook';
import { buildWorkbook, worksheet } from '../read/test-utils/package-builder';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import { writeWorkbook } from './index';
import {
	childNames,
	openPackage,
	partRoot,
	pick,
	pickAll,
	validatePackage,
} from './test-utils/package-check';

const SHEET = 'xl/worksheets/sheet1.xml';

async function source(before: string, after: string): Promise<Uint8Array> {
	return buildWorkbook({
		sheets: [
			{ name: 'S', xml: worksheet('<row r="1"><c r="A1"><v>1</v></c></row>', { before, after }) },
		],
	});
}

const PRINT_XML =
	'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
	'<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
	'<headerFooter scaleWithDoc="0" alignWithMargins="0"><oddFooter>Page &amp;P</oddFooter></headerFooter>' +
	'<rowBreaks count="3" manualBreakCount="2"><brk id="5" max="16383" man="1"/><brk id="12" max="16383"/>' +
	'<brk id="20" min="2" max="10" man="1"/></rowBreaks>' +
	'<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>';

describe('page breaks and fit to page: read', () => {
	it('reads manual breaks, fitToPage and header / footer flags', async () => {
		const wb = await readWorkbook(
			await source('<sheetPr><pageSetUpPr autoPageBreaks="0" fitToPage="1"/></sheetPr>', PRINT_XML),
		);
		const sheet = wb.sheets[0];
		expect(sheet.rowBreaks).toEqual([5, 20]);
		expect(sheet.colBreaks).toEqual([3]);
		expect(sheet.pageSetup).toEqual({
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 0,
			fitToPage: true,
		});
		expect(sheet.headerFooter).toEqual({
			scaleWithDoc: false,
			alignWithMargins: false,
			oddFooter: 'Page &P',
		});
	});

	it('leaves the fields unset when the file has none', async () => {
		const sheet = (await readWorkbook(await source('', ''))).sheets[0];
		expect(sheet.rowBreaks).toBeUndefined();
		expect(sheet.colBreaks).toBeUndefined();
		expect(sheet.pageSetup).toBeUndefined();
	});
});

describe('page breaks and fit to page: write', () => {
	it('writes a sheet created in memory', async () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		putCell(sheet, 0, 0, { value: 1 });
		sheet.rowBreaks = [9, 4];
		sheet.colBreaks = [2];
		sheet.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
		const pkg = await openPackage(await writeWorkbook(wb));
		expect(await validatePackage(pkg)).toEqual([]);
		const root = await partRoot(pkg, SHEET);
		const names = childNames(root);
		expect(names.indexOf('pageSetup')).toBeLessThan(names.indexOf('rowBreaks'));
		expect(names.indexOf('rowBreaks')).toBeLessThan(names.indexOf('colBreaks'));
		expect(pick(root, 'sheetPr', 'pageSetUpPr')?.attrs).toEqual({ fitToPage: '1' });
		expect(pick(root, 'rowBreaks')?.attrs).toEqual({ count: '2', manualBreakCount: '2' });
		expect(pickAll(pick(root, 'rowBreaks'), 'brk').map((b) => b.attrs)).toEqual([
			{ id: '4', max: '16383', man: '1' },
			{ id: '9', max: '16383', man: '1' },
		]);
		expect(pickAll(pick(root, 'colBreaks'), 'brk').map((b) => b.attrs)).toEqual([
			{ id: '2', max: '1048575', man: '1' },
		]);
	});

	it('round-trips through read and write', async () => {
		const wb = await readWorkbook(
			await source('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', PRINT_XML),
		);
		wb.sheets[0].dirty = true;
		const again = await readWorkbook(await writeWorkbook(wb));
		const sheet = again.sheets[0];
		expect(sheet.rowBreaks).toEqual([5, 20]);
		expect(sheet.colBreaks).toEqual([3]);
		expect(sheet.pageSetup?.fitToPage).toBe(true);
		expect(sheet.headerFooter?.scaleWithDoc).toBe(false);
	});

	it('keeps a clean sheet byte-for-byte', async () => {
		const bytes = await source('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', PRINT_XML);
		const wb = await readWorkbook(bytes);
		const out = await writeWorkbook(wb);
		const [a, b] = [await openPackage(bytes), await openPackage(out)];
		expect(await b.text(SHEET)).toBe(await a.text(SHEET));
	});

	it('edits breaks: original attributes and automatic breaks survive', async () => {
		const wb = await readWorkbook(
			await source(
				'<sheetPr codeName="X"><pageSetUpPr autoPageBreaks="0" fitToPage="1"/></sheetPr>',
				PRINT_XML,
			),
		);
		const sheet = wb.sheets[0];
		sheet.rowBreaks = [20, 30];
		sheet.colBreaks = [];
		sheet.pageSetup = { ...sheet.pageSetup, fitToPage: false };
		sheet.dirty = true;
		const root = await partRoot(await openPackage(await writeWorkbook(wb)), SHEET);
		expect(pick(root, 'rowBreaks')?.attrs).toEqual({ count: '3', manualBreakCount: '2' });
		expect(pickAll(pick(root, 'rowBreaks'), 'brk').map((b) => b.attrs)).toEqual([
			{ id: '12', max: '16383' },
			{ id: '20', min: '2', max: '10', man: '1' },
			{ id: '30', max: '16383', man: '1' },
		]);
		expect(pick(root, 'colBreaks')).toBeUndefined();
		expect(pick(root, 'sheetPr')?.attrs).toEqual({ codeName: 'X' });
		expect(pick(root, 'sheetPr', 'pageSetUpPr')?.attrs).toEqual({ autoPageBreaks: '0' });
	});

	it('keeps the original elements while the model leaves them unset', async () => {
		const wb = await readWorkbook(await source('', PRINT_XML));
		const sheet = wb.sheets[0];
		sheet.rowBreaks = undefined;
		sheet.colBreaks = undefined;
		sheet.dirty = true;
		const root = await partRoot(await openPackage(await writeWorkbook(wb)), SHEET);
		expect(pickAll(pick(root, 'rowBreaks'), 'brk')).toHaveLength(3);
		expect(pick(root, 'colBreaks')).toBeDefined();
	});
});

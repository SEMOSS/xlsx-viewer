import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { readWorkbook } from '../read';
import { createWorkbook, nextSheetId } from '../sheet/workbook';
import { createWorksheet, putCell } from '../sheet/worksheet';
import { StylePool } from '../styles/style-pool';
import { writeWorkbook } from './index';
import { fixtureBytes, fixtureNames, withCalcChain } from './test-utils/fixtures';
import { openPackage, partRoot, pick, pickAll, validatePackage } from './test-utils/package-check';

const HR = 'issue-132-hr-deck--Workbook1.xlsx';
const CALLOUTS = 'chart-label-callouts--Microsoft_Excel_Worksheet.xlsx';
const NOW = new Date('2025-02-03T04:05:06Z');

async function samePart(a: Uint8Array, b: Uint8Array, path: string): Promise<boolean> {
	const [pa, pb] = [await openPackage(a), await openPackage(b)];
	const [x, y] = [await pa.bytes(path), await pb.bytes(path)];
	return !!x && !!y && Buffer.compare(Buffer.from(x), Buffer.from(y)) === 0;
}

describe('read -> write without changes', () => {
	it.each(fixtureNames())('%s keeps every part byte-identical', async (name) => {
		const bytes = fixtureBytes(name);
		const src = await openPackage(bytes);
		const out = await openPackage(await writeWorkbook(await readWorkbook(bytes)));
		expect(out.partPaths().sort()).toEqual(src.partPaths().sort());
		for (const p of src.partPaths()) {
			if (p !== '[Content_Types].xml') {
				expect(await out.bytes(p), p).toEqual(await src.bytes(p));
			}
		}
	});
});

describe('read -> modify -> write -> read', () => {
	it('edits cells: only the sheet and core properties change, calc chain is dropped', async () => {
		const bytes = await withCalcChain(fixtureBytes(HR));
		const wb = await readWorkbook(bytes);
		const sheet = wb.sheets[0];
		const style = sheet.cells.get(cellKey(1, 1))?.style;
		putCell(sheet, 1, 1, { value: 0.5, style });
		putCell(sheet, 7, 0, { value: 'New text' });
		putCell(sheet, 7, 1, {
			value: 'Bold tail',
			richText: [{ text: 'Bold', font: { bold: true } }, { text: ' tail' }],
		});
		const out = await writeWorkbook(wb, { now: NOW });
		const pkg = await openPackage(out);
		expect(await validatePackage(pkg)).toEqual([]);

		expect(pkg.has('xl/calcChain.xml')).toBe(false);
		expect(pkg.contentTypes.overrides.has('xl/calcChain.xml')).toBe(false);
		expect((await pkg.rels('xl/workbook.xml')).some((r) => r.type.endsWith('/calcChain'))).toBe(
			false,
		);
		for (const p of [
			'xl/workbook.xml',
			'xl/styles.xml',
			'xl/theme/theme1.xml',
			'xl/tables/table1.xml',
			'docProps/app.xml',
		]) {
			expect(await samePart(bytes, out, p), p).toBe(true);
		}
		const core = await partRoot(pkg, 'docProps/core.xml');
		expect(pick(core, 'modified')?.children[0]).toBe('2025-02-03T04:05:06Z');
		expect(pick(core, 'creator')?.children[0]).toBe('kingpub');

		const sst = pickAll(await partRoot(pkg, 'xl/sharedStrings.xml'), 'si');
		expect(sst).toHaveLength(9);
		expect(pick(sst[1], 't')?.children[0]).toBe('列1');
		const sheetXml = await partRoot(pkg, 'xl/worksheets/sheet1.xml');
		expect(pick(sheetXml, 'pageSetup')?.attrs).toMatchObject({ paperSize: '9', 'r:id': 'rId1' });
		expect(pick(sheetXml, 'phoneticPr')).toBeDefined();
		expect(pick(sheetXml, 'tableParts', 'tablePart')?.attrs['r:id']).toBe('rId2');

		const back = await readWorkbook(out);
		const s = back.sheets[0];
		expect(s.cells.get(cellKey(1, 1))).toMatchObject({ value: 0.5, style });
		expect(s.cells.get(cellKey(5, 1))?.value).toBeCloseTo(0.07, 12);
		expect(s.cells.get(cellKey(0, 0))?.value).toBe(' ');
		expect(s.cells.get(cellKey(7, 0))?.value).toBe('New text');
		expect(s.cells.get(cellKey(7, 1))?.richText?.[0]).toMatchObject({
			text: 'Bold',
			font: { bold: true },
		});
		expect(s.tables.map((t) => t.name)).toEqual(['表1']);
		expect(s.view.zoomScale).toBe(160);
	});

	it('adds a styled sheet, renames and reorders sheets', async () => {
		const bytes = fixtureBytes(HR);
		const wb = await readWorkbook(bytes);
		const pool = new StylePool(wb.styles);
		const bold = pool.derive(0, { font: { bold: true }, numFmt: '0.0%' });
		const added = createWorksheet('Added', nextSheetId(wb));
		putCell(added, 0, 0, { value: 'Header', style: bold });
		putCell(added, 1, 0, { value: 2, formula: 'Data!B2*50+IFNA(1,0)' });
		wb.sheets[0].name = 'Data';
		wb.sheets.unshift(added);
		wb.view.activeTab = 1;
		wb.dirty = true;
		const out = await writeWorkbook(wb, { now: NOW });
		const pkg = await openPackage(out);
		expect(await validatePackage(pkg)).toEqual([]);
		expect(await samePart(bytes, out, 'xl/worksheets/sheet1.xml')).toBe(true);
		expect(pkg.has('xl/worksheets/sheet2.xml')).toBe(true);

		const wbXml = await partRoot(pkg, 'xl/workbook.xml');
		expect(pickAll(pick(wbXml, 'sheets'), 'sheet').map((s) => s.attrs.name)).toEqual([
			'Added',
			'Data',
		]);
		expect(pick(wbXml, 'bookViews', 'workbookView')?.attrs.activeTab).toBe('1');
		expect(pick(wbXml, 'extLst')).toBeDefined();
		const app = await partRoot(pkg, 'docProps/app.xml');
		expect(
			pickAll(pick(app, 'TitlesOfParts', 'vector'), 'lpstr').map((t) => t.children[0]),
		).toEqual(['Added', 'Data']);

		const styles = await partRoot(pkg, 'xl/styles.xml');
		const fonts = pickAll(pick(styles, 'fonts'), 'font');
		expect(fonts).toHaveLength(3);
		expect(pick(fonts[0], 'name')?.attrs.val).toBe('宋体');
		expect(pick(styles, 'fonts')?.attrs['x14ac:knownFonts']).toBe('1');

		const back = await readWorkbook(out);
		expect(back.sheets.map((s) => s.name)).toEqual(['Added', 'Data']);
		const head = back.sheets[0].cells.get(0)!;
		expect(back.styles.fonts[back.styles.cellXfs[head.style!].fontId].bold).toBe(true);
		expect(back.styles.numFmts[back.styles.cellXfs[head.style!].numFmtId]).toBe('0.0%');
		expect(back.sheets[0].cells.get(cellKey(1, 0))?.formula).toBe('Data!B2*50+IFNA(1,0)');
		expect(await pkg.text('xl/worksheets/sheet2.xml')).toContain('_xlfn.IFNA(1,0)');
		expect(back.sheets[1].cells.get(cellKey(1, 0))?.value).toBe('硕士');
	});

	it('deletes a sheet together with the parts only it used', async () => {
		const wb = await readWorkbook(fixtureBytes(HR));
		const only = createWorksheet('Only', 7);
		putCell(only, 0, 0, { value: 1 });
		wb.sheets = [only];
		wb.view.activeTab = 0;
		wb.dirty = true;
		const pkg = await openPackage(await writeWorkbook(wb));
		expect(await validatePackage(pkg)).toEqual([]);
		for (const gone of [
			'xl/worksheets/sheet1.xml',
			'xl/tables/table1.xml',
			'xl/printerSettings/printerSettings1.bin',
		]) {
			expect(pkg.has(gone), gone).toBe(false);
			expect(pkg.contentTypes.overrides.has(gone)).toBe(false);
		}
		const sheets = pickAll(pick(await partRoot(pkg, 'xl/workbook.xml'), 'sheets'), 'sheet');
		expect(sheets.map((s) => [s.attrs.name, s.attrs.sheetId])).toEqual([['Only', '7']]);
		expect(
			(await readWorkbook(await writeWorkbook(await readWorkbook(fixtureBytes(HR))))).sheets,
		).toHaveLength(1);
	});

	it('duplicates a loaded sheet into fresh parts', async () => {
		const wb = await readWorkbook(fixtureBytes(CALLOUTS));
		const original = wb.sheets[0];
		const copy = {
			...original,
			id: 'copy',
			name: 'Copy',
			sheetId: 2,
			cells: new Map(original.cells),
			tables: original.tables.map((t) => ({
				...t,
				name: 'Table1',
				displayName: 'Table1',
				dirty: true,
			})),
			dirty: true,
		};
		wb.sheets.push(copy);
		wb.dirty = true;
		const out = await writeWorkbook(wb);
		const pkg = await openPackage(out);
		expect(await validatePackage(pkg)).toEqual([]);
		const back = await readWorkbook(out);
		expect(back.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Copy']);
		expect(back.sheets[1].cells.get(cellKey(1, 1))?.value).toBe(4.3);
		const names = back.sheets.flatMap((s) => s.tables.map((t) => t.name));
		expect(new Set(names).size).toBe(2);
		const ids = back.sheets.flatMap((s) => s.tables.map((t) => t.id));
		expect(new Set(ids).size).toBe(2);
	});
});

describe('package-level bookkeeping', () => {
	it('drops the calc chain when a sheet is deleted, even if no sheet is rewritten', async () => {
		const wb = createWorkbook({ sheetNames: ['Keep', 'Drop'] });
		putCell(wb.sheets[0], 0, 0, { value: 1, formula: '1+1' });
		putCell(wb.sheets[1], 0, 0, { value: 2, formula: '1+1' });
		const bytes = await withCalcChain(await writeWorkbook(wb));
		const loaded = await readWorkbook(bytes);
		loaded.sheets.pop();
		loaded.dirty = true;
		const out = await writeWorkbook(loaded);
		const pkg = await openPackage(out);
		expect(await validatePackage(pkg)).toEqual([]);
		expect(pkg.has('xl/calcChain.xml')).toBe(false);
		expect(pkg.has('xl/worksheets/sheet2.xml')).toBe(false);
		expect(await samePart(bytes, out, 'xl/worksheets/sheet1.xml')).toBe(true);
	});

	it('does not add document properties to an unchanged package that has none', async () => {
		const zip = await JSZip.loadAsync(fixtureBytes(HR));
		zip.remove('docProps/core.xml');
		zip.remove('docProps/app.xml');
		const rels = await zip.file('_rels/.rels')!.async('string');
		zip.file('_rels/.rels', rels.replace(/<Relationship [^>]*docProps[^>]*\/>/g, ''));
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const src = await openPackage(bytes);
		const out = await openPackage(await writeWorkbook(await readWorkbook(bytes)));
		expect(out.partPaths().sort()).toEqual(src.partPaths().sort());
		expect(await out.bytes('_rels/.rels')).toEqual(await src.bytes('_rels/.rels'));
	});
});

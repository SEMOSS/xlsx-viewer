import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { decryptPackage, encryptionInfoOf, isEncryptedPackageFile } from '../package';
import { readWorkbook, XlsxPasswordError } from '../read';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import type { XlsxWarning } from '../types/warnings';
import { parseXml } from '../xml';
import { buildAppXml, buildCoreXml } from './doc-props';
import { writeWorkbook } from './index';
import { buildThemeXml } from './theme-xml';
import { fixtureBytes } from './test-utils/fixtures';
import {
	childNames,
	openPackage,
	partRoot,
	pick,
	pickAll,
	validatePackage,
} from './test-utils/package-check';

const NOW = new Date('2024-06-07T08:09:10.500Z');

describe('in-memory workbooks', () => {
	it('produce a complete, valid package', async () => {
		const wb = createWorkbook({ sheetNames: ['One', 'Two'], title: 'Report' });
		putCell(wb.sheets[0], 0, 0, { value: 'x' });
		const bytes = await writeWorkbook(wb, { now: NOW });
		const pkg = await openPackage(bytes);
		expect(pkg.partPaths().sort()).toEqual(
			[
				'[Content_Types].xml',
				'_rels/.rels',
				'docProps/app.xml',
				'docProps/core.xml',
				'xl/_rels/workbook.xml.rels',
				'xl/sharedStrings.xml',
				'xl/styles.xml',
				'xl/theme/theme1.xml',
				'xl/workbook.xml',
				'xl/worksheets/sheet1.xml',
				'xl/worksheets/sheet2.xml',
			].sort(),
		);
		expect(await validatePackage(pkg)).toEqual([]);
		expect(Object.keys((await JSZip.loadAsync(bytes)).files)[0]).toBe('[Content_Types].xml');
		expect(pkg.contentTypes.get('xl/workbook.xml')).toBe(
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
		);
		expect((await pkg.rels('')).map((r) => r.type.split('/').pop())).toEqual([
			'officeDocument',
			'core-properties',
			'extended-properties',
		]);
		const core = await partRoot(pkg, 'docProps/core.xml');
		expect(pick(core, 'title')?.children[0]).toBe('Report');
		expect(pick(core, 'modified')?.children[0]).toBe('2024-06-07T08:09:10Z');
		const sheet = await partRoot(pkg, 'xl/worksheets/sheet1.xml');
		expect(pick(sheet, 'sheetViews', 'sheetView')?.attrs.tabSelected).toBe('1');
		expect(
			pick(await partRoot(pkg, 'xl/worksheets/sheet2.xml'), 'sheetViews', 'sheetView')?.attrs
				.tabSelected,
		).toBeUndefined();
		const back = await readWorkbook(bytes);
		expect(back.sheets.map((s) => s.name)).toEqual(['One', 'Two']);
		expect(back.properties.title).toBe('Report');
		expect(back.theme.colors.accent1).toBe('4472C4');
	});

	it('encrypts with a password that reading then requires', async () => {
		const wb = createWorkbook({ sheetNames: ['Secret'] });
		putCell(wb.sheets[0], 0, 0, { value: 'classified' });
		const bytes = await writeWorkbook(wb, { password: 'Pa55word', now: NOW });
		expect(isEncryptedPackageFile(bytes)).toBe(true);
		expect(encryptionInfoOf(bytes)).toMatchObject({
			type: 'agile',
			keyData: { keyBits: 256, hashAlgorithm: 'SHA512' },
			passwordKey: { spinCount: 100000 },
		});
		const missing = await readWorkbook(bytes).catch((e: unknown) => e);
		expect(missing).toBeInstanceOf(XlsxPasswordError);
		expect((missing as XlsxPasswordError).reason).toBe('required');
		await expect(readWorkbook(bytes, { password: 'pa55word' })).rejects.toMatchObject({
			reason: 'incorrect',
		});
		const back = await readWorkbook(bytes, { password: 'Pa55word' });
		expect(back.encrypted).toBe(true);
		expect(back.sheets[0].cells.get(cellKey(0, 0))?.value).toBe('classified');
		// Saving without the password writes a plain package again.
		const plain = await writeWorkbook(back, { now: NOW });
		expect(isEncryptedPackageFile(plain)).toBe(false);
		expect((await readWorkbook(plain)).encrypted).toBeUndefined();
	});

	it('round-trips a real workbook through encryption unchanged', async () => {
		const source = await readWorkbook(fixtureBytes('issue-132-hr-deck--Workbook1.xlsx'));
		const plain = await openPackage(await writeWorkbook(source, { now: NOW }));
		const encrypted = await writeWorkbook(source, { password: 'x', now: NOW });
		const decrypted = await openPackage(await decryptPackage(encrypted, 'x'));
		expect(decrypted.partPaths().sort()).toEqual(plain.partPaths().sort());
		for (const part of plain.partPaths()) {
			expect(await decrypted.text(part), part).toBe(await plain.text(part));
		}
		const back = await readWorkbook(encrypted, { password: 'x' });
		expect(back.sheets.map((s) => s.name)).toEqual(source.sheets.map((s) => s.name));
	});

	it('fixes sheet names, ids, visibility and defined-name scopes that Excel would repair', async () => {
		const wb = createWorkbook({ sheetNames: ['Data', 'data', 'Bad:Name?'] });
		wb.sheets[1].sheetId = 1;
		for (const s of wb.sheets) {
			s.state = 'hidden';
		}
		wb.view.activeTab = 2;
		wb.definedNames = [
			{ name: 'A', formula: 'Data!$A$1' },
			{ name: 'a', formula: 'Data!$A$2' },
			{ name: 'Scoped', formula: '1', localSheet: 9 },
		];
		const warnings: XlsxWarning[] = [];
		const pkg = await openPackage(await writeWorkbook(wb, { onWarning: (w) => warnings.push(w) }));
		expect(await validatePackage(pkg)).toEqual([]);
		const root = await partRoot(pkg, 'xl/workbook.xml');
		const sheets = pickAll(pick(root, 'sheets'), 'sheet');
		expect(sheets.map((s) => [s.attrs.name, s.attrs.sheetId, s.attrs.state])).toEqual([
			['Data', '1', undefined],
			['data (2)', '4', 'hidden'],
			['Bad_Name_', '3', 'hidden'],
		]);
		expect(pickAll(pick(root, 'definedNames'), 'definedName').map((d) => d.attrs.name)).toEqual([
			'A',
		]);
		expect(warnings.filter((w) => w.code === 'sheet-renamed')).toHaveLength(2);
	});

	it('asks Excel to recalculate when formulas are stale or were never calculated', async () => {
		const flag = async (cell: Parameters<typeof putCell>[3]) => {
			const wb = createWorkbook();
			putCell(wb.sheets[0], 0, 0, cell);
			const root = await partRoot(await openPackage(await writeWorkbook(wb)), 'xl/workbook.xml');
			return pick(root, 'calcPr')?.attrs;
		};
		expect(await flag({ value: 1, formula: 'A2+1' })).toEqual({ calcId: '191029' });
		expect(await flag({ value: 1, formula: 'A2+1', stale: true })).toMatchObject({
			fullCalcOnLoad: '1',
		});
		expect(await flag({ value: null, formula: 'A2+1' })).toMatchObject({ fullCalcOnLoad: '1' });
	});

	it('reports problems through onWarning', async () => {
		const wb = createWorkbook();
		putCell(wb.sheets[0], 0, 0, { value: 1, style: 42 });
		wb.sheets[0].merges = [
			{ top: 0, left: 0, bottom: 1, right: 1 },
			{ top: 1, left: 1, bottom: 2, right: 2 },
		];
		const warnings: XlsxWarning[] = [];
		await writeWorkbook(wb, { onWarning: (w) => warnings.push(w) });
		expect(warnings.map((w) => w.code).sort()).toEqual(['invalid-style', 'merge-dropped']);
	});

	it('saves 100k cells within the performance budget', async () => {
		const wb = createWorkbook();
		const s = wb.sheets[0];
		for (let r = 0; r < 2000; r++) {
			for (let c = 0; c < 50; c++) {
				s.cells.set(
					cellKey(r, c),
					c % 7 === 0
						? { value: r, formula: `SUM(A${r + 1}:C${r + 1})*2` }
						: c % 5 === 0
							? { value: `s${r % 300}` }
							: { value: r * c + 0.5 },
				);
			}
		}
		await writeWorkbook(wb);
		const start = performance.now();
		await writeWorkbook(wb);
		expect(performance.now() - start).toBeLessThan(500);
	});
});

describe('loaded workbooks', () => {
	it('keeps parts that were unreachable in the source', async () => {
		const zip = await JSZip.loadAsync(fixtureBytes('ole-editable--Microsoft_Excel_Worksheet.xlsx'));
		zip.file('customXml/stray.xml', '<x/>');
		const wb = await readWorkbook(await zip.generateAsync({ type: 'uint8array' }));
		wb.sheets[0].dirty = true;
		const pkg = await openPackage(await writeWorkbook(wb));
		expect(pkg.has('customXml/stray.xml')).toBe(true);
	});

	it('keeps Company and Manager edits when the file already has app.xml', async () => {
		const wb = await readWorkbook(fixtureBytes('ole-editable--Microsoft_Excel_Worksheet.xlsx'));
		wb.properties = { ...wb.properties, company: 'Contoso Ltd', manager: 'Dana' };
		wb.dirty = true;
		const back = await readWorkbook(await writeWorkbook(wb));
		expect(back.properties.company).toBe('Contoso Ltd');
		expect(back.properties.manager).toBe('Dana');
	});

	it('writes a theme from the model when its raw XML was dropped', async () => {
		const wb = await readWorkbook(fixtureBytes('ole-editable--Microsoft_Excel_Worksheet.xlsx'));
		wb.theme = {
			...wb.theme,
			rawXml: undefined,
			colors: { ...wb.theme.colors, accent1: '123456' },
		};
		wb.dirty = true;
		const back = await readWorkbook(await writeWorkbook(wb));
		expect(back.theme.colors.accent1).toBe('123456');
	});
});

describe('theme and document properties', () => {
	it('generates a complete theme from the model', () => {
		const wb = createWorkbook();
		wb.theme.colors.dk1 = '111111';
		wb.theme.fonts.majorLatin = 'Aptos Display';
		const root = parseXml(buildThemeXml(wb.theme)).root;
		const elements = pick(root, 'themeElements')!;
		expect(childNames(elements)).toEqual(['clrScheme', 'fontScheme', 'fmtScheme']);
		expect(childNames(pick(elements, 'clrScheme'))).toHaveLength(12);
		expect(pick(elements, 'clrScheme', 'dk1', 'srgbClr')?.attrs.val).toBe('111111');
		expect(pick(elements, 'clrScheme', 'lt1', 'sysClr')?.attrs).toEqual({
			val: 'window',
			lastClr: 'FFFFFF',
		});
		expect(pick(elements, 'fontScheme', 'majorFont', 'latin')?.attrs.typeface).toBe(
			'Aptos Display',
		);
		expect(childNames(pick(elements, 'fmtScheme'))).toEqual([
			'fillStyleLst',
			'lnStyleLst',
			'effectStyleLst',
			'bgFillStyleLst',
		]);
		expect(childNames(root)).toEqual(['themeElements', 'objectDefaults', 'extraClrSchemeLst']);
	});

	it('updates core properties on top of the original part', () => {
		const original = parseXml(
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Ann</dc:creator><cp:revision>4</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">2001-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>',
		).root;
		const root = parseXml(
			buildCoreXml(
				{ creator: 'Ann', title: 'T', created: '2020-02-02T02:02:02.123Z' },
				original,
				NOW,
			),
		).root;
		expect(childNames(root)).toEqual(['title', 'creator', 'revision', 'created', 'modified']);
		expect(pick(root, 'created')?.children[0]).toBe('2020-02-02T02:02:02Z');
		expect(pick(root, 'modified')?.children[0]).toBe('2024-06-07T08:09:10Z');
	});

	it('replaces only the worksheet group of the app titles', () => {
		const original = parseXml(
			'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Arbeitsblätter</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Benannte Bereiche</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>A</vt:lpstr><vt:lpstr>B</vt:lpstr><vt:lpstr>A!Print_Area</vt:lpstr></vt:vector></TitlesOfParts><AppVersion>16.0300</AppVersion></Properties>',
		).root;
		const root = parseXml(
			buildAppXml(
				{ sheetNames: ['A', 'B', 'C'], previousSheetNames: ['A', 'B'], application: 'x' },
				original,
			),
		).root;
		const pairs = pickAll(pick(root, 'HeadingPairs', 'vector'), 'variant').map(
			(v) => pick(v, 'lpstr')?.children[0] ?? pick(v, 'i4')?.children[0],
		);
		expect(pairs).toEqual(['Arbeitsblätter', '3', 'Benannte Bereiche', '1']);
		expect(
			pickAll(pick(root, 'TitlesOfParts', 'vector'), 'lpstr').map((t) => t.children[0]),
		).toEqual(['A', 'B', 'C', 'A!Print_Area']);
		expect(pick(root, 'TitlesOfParts', 'vector')?.attrs.size).toBe('4');
		expect(pick(root, 'Application')?.children[0]).toBe('Microsoft Excel');
	});

	it('writes custom document properties with inferred types', async () => {
		const wb = createWorkbook();
		wb.properties.custom = [
			{ name: 'Reviewed', value: true },
			{ name: 'Count', value: 3 },
			{ name: 'Ratio', value: 0.5 },
			{ name: 'Owner', value: 'Ann' },
		];
		const pkg = await openPackage(await writeWorkbook(wb));
		expect(await validatePackage(pkg)).toEqual([]);
		const root = await partRoot(pkg, 'docProps/custom.xml');
		expect(
			pickAll(root, 'property').map((p) => [p.attrs.pid, p.attrs.name, childNames(p)[0]]),
		).toEqual([
			['2', 'Reviewed', 'bool'],
			['3', 'Count', 'i4'],
			['4', 'Ratio', 'r8'],
			['5', 'Owner', 'lpwstr'],
		]);
		expect((await pkg.rels('')).some((r) => r.type.endsWith('/custom-properties'))).toBe(true);
	});
});

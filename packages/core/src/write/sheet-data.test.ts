import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { createWorksheet } from '../sheet/worksheet';
import type { Cell } from '../types/cell';
import { ERRORS } from '../values/errors';
import { parseXml } from '../xml';
import { CellMetadata } from './metadata';
import { SharedStringTable } from './shared-strings';
import { buildSheetData, type SheetDataContext } from './sheet-data';
import { childNames, pick, pickAll } from './test-utils/package-check';
import { formatXmlNumber } from './xml-utils';

function ctx(styleCount = 10): SheetDataContext {
	return { sst: SharedStringTable.empty(), styleCount, metadata: new CellMetadata() };
}

function sheetWith(cells: Array<[number, number, Cell]>) {
	const sheet = createWorksheet('S', 1);
	for (const [r, c, cell] of cells) {
		sheet.cells.set(cellKey(r, c), cell);
	}
	return sheet;
}

function cellsOf(xml: string) {
	const root = parseXml(xml).root;
	return pickAll(root, 'row').flatMap((row) => pickAll(row, 'c'));
}

describe('number formatting', () => {
	it.each([
		[0, '0'],
		[-0, '0'],
		[42, '42'],
		[-1.25, '-1.25'],
		[0.1 + 0.2, '0.30000000000000004'],
		[0.07, '0.07'],
		[1.5e-7, '1.5E-7'],
		[1e21, '1E+21'],
		[123456789012345680000, '123456789012345680000'],
		[Number.MAX_VALUE, '1.7976931348623157E+308'],
	])('%s -> %s', (n, s) => {
		expect(formatXmlNumber(n)).toBe(s);
		expect(Number(s)).toBe(n === 0 ? 0 : n);
	});
});

describe('buildSheetData', () => {
	it('writes values with the right types', () => {
		const c = ctx();
		const sheet = sheetWith([
			[0, 0, { value: 'text' }],
			[0, 1, { value: 3.5 }],
			[0, 2, { value: true }],
			[0, 3, { value: false }],
			[0, 4, { value: ERRORS.DIV0 }],
			[0, 5, { value: 'text' }],
		]);
		const res = buildSheetData(sheet, c);
		const cells = cellsOf(res.xml);
		expect(cells.map((x) => [x.attrs.r, x.attrs.t, pick(x, 'v')?.children[0]])).toEqual([
			['A1', 's', '0'],
			['B1', undefined, '3.5'],
			['C1', 'b', '1'],
			['D1', 'b', '0'],
			['E1', 'e', '#DIV/0!'],
			['F1', 's', '0'],
		]);
		expect(c.sst.size).toBe(1);
		expect(c.sst.refs).toBe(2);
		expect(res.used).toEqual({ top: 0, left: 0, bottom: 0, right: 5 });
	});

	it('orders rows and cells numerically regardless of insertion order', () => {
		const sheet = sheetWith([
			[10, 2, { value: 1 }],
			[1, 27, { value: 2 }],
			[1, 3, { value: 3 }],
			[0, 0, { value: 4 }],
		]);
		const refs = cellsOf(buildSheetData(sheet, ctx()).xml).map((x) => x.attrs.r);
		expect(refs).toEqual(['A1', 'D2', 'AB2', 'C11']);
	});

	it('writes formulas with prefixes and typed cached values', () => {
		const sheet = sheetWith([
			[0, 0, { value: 3, formula: 'SUM(B1:B2)' }],
			[1, 0, { value: 'x', formula: 'CONCAT("a","b")&IFS(TRUE,"x")' }],
			[2, 0, { value: true, formula: 'XOR(TRUE,FALSE)' }],
			[3, 0, { value: ERRORS.NA, formula: 'XLOOKUP(1,B:B,C:C)' }],
			[4, 0, { value: null, formula: 'A1*2' }],
			[5, 0, { value: 'a<b&c', formula: '"a<b&c"' }],
		]);
		const cells = cellsOf(buildSheetData(sheet, ctx()).xml);
		const f = (i: number) => pick(cells[i], 'f')?.children[0];
		expect(f(0)).toBe('SUM(B1:B2)');
		expect(f(1)).toBe('_xlfn.CONCAT("a","b")&_xlfn.IFS(TRUE,"x")');
		expect(cells[1].attrs.t).toBe('str');
		expect(cells[2].attrs.t).toBe('b');
		expect(f(3)).toBe('_xlfn.XLOOKUP(1,B:B,C:C)');
		expect(cells[3].attrs.t).toBe('e');
		expect(pick(cells[4], 'v')).toBeUndefined();
		expect(pick(cells[5], 'v')?.children[0]).toBe('a<b&c');
		expect(childNames(cells[0])).toEqual(['f', 'v']);
	});

	it('writes legacy array formulas and values of array members', () => {
		const sheet = sheetWith([
			[
				0,
				0,
				{
					value: 1,
					formula: 'A3:B4*2',
					array: { ref: { top: 0, left: 0, bottom: 1, right: 1 }, dynamic: false },
				},
			],
			[0, 1, { value: 2, arrayAnchor: cellKey(0, 0), formula: 'A3:B4*2' }],
			[1, 0, { value: 3, arrayAnchor: cellKey(0, 0) }],
		]);
		const c = ctx();
		const cells = cellsOf(buildSheetData(sheet, c).xml);
		expect(pick(cells[0], 'f')?.attrs).toEqual({ t: 'array', ref: 'A1:B2' });
		expect(cells[0].attrs.cm).toBeUndefined();
		expect(pick(cells[1], 'f')).toBeUndefined();
		expect(pick(cells[1], 'v')?.children[0]).toBe('2');
		expect(c.metadata.changed).toBe(false);
	});

	it('links dynamic arrays to an XLDAPR cell metadata record', () => {
		const sheet = sheetWith([
			[
				0,
				0,
				{
					value: 1,
					formula: 'SEQUENCE(3)',
					array: { ref: { top: 0, left: 0, bottom: 2, right: 0 }, dynamic: true },
				},
			],
			[1, 0, { value: 2, arrayAnchor: cellKey(0, 0) }],
			[
				0,
				2,
				{
					value: 'a',
					formula: 'UNIQUE(B1:B9)',
					array: { ref: { top: 0, left: 2, bottom: 0, right: 2 }, dynamic: true },
				},
			],
		]);
		const c = ctx();
		const cells = cellsOf(buildSheetData(sheet, c).xml);
		expect(cells[0].attrs.cm).toBe('1');
		expect(pick(cells[0], 'f')?.children[0]).toBe('_xlfn.SEQUENCE(3)');
		expect(cells[1].attrs.r).toBe('C1');
		expect(cells[1].attrs.cm).toBe('1');
		expect(cells[2].attrs.cm).toBeUndefined();
		expect(c.metadata.changed).toBe(true);
		const md = parseXml(c.metadata.serialize()).root;
		expect(childNames(md)).toEqual(['metadataTypes', 'futureMetadata', 'cellMetadata']);
		expect(pick(md, 'metadataTypes', 'metadataType')?.attrs.name).toBe('XLDAPR');
		expect(pick(md, 'cellMetadata')?.attrs.count).toBe('1');
	});

	it('keeps valid cm / vm / ph and drops dangling metadata references', () => {
		const md = CellMetadata.parse(
			'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="1"><metadataType name="XLRICHVALUE"/></metadataTypes><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata><valueMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk></valueMetadata></metadata>',
			'xl/metadata.xml',
		);
		const sheet = sheetWith([
			[0, 0, { value: 1, cm: 1, vm: 2, ph: true }],
			[0, 1, { value: 1, cm: 5, vm: 9 }],
		]);
		const cells = cellsOf(buildSheetData(sheet, { ...ctx(), metadata: md }).xml);
		expect(cells[0].attrs).toMatchObject({ cm: '1', vm: '2', ph: '1' });
		expect(cells[1].attrs.cm).toBeUndefined();
		expect(cells[1].attrs.vm).toBeUndefined();
	});

	it('writes style-only cells, validates style indices and skips empty cells', () => {
		const sheet = sheetWith([
			[0, 0, { value: null, style: 3 }],
			[0, 1, { value: null }],
			[0, 2, { value: 5, style: 99 }],
		]);
		const res = buildSheetData(sheet, ctx(4));
		const cells = cellsOf(res.xml);
		expect(cells.map((x) => x.attrs)).toEqual([{ r: 'A1', s: '3' }, { r: 'C1' }]);
		expect(res.invalidStyles).toBe(1);
	});

	it('writes non-finite numbers as #NUM!', () => {
		const res = buildSheetData(
			sheetWith([
				[0, 0, { value: Number.NaN }],
				[0, 1, { value: Infinity, formula: '1/0' }],
			]),
			ctx(),
		);
		const cells = cellsOf(res.xml);
		expect(cells.map((x) => [x.attrs.t, pick(x, 'v')?.children[0]])).toEqual([
			['e', '#NUM!'],
			['e', '#NUM!'],
		]);
		expect(res.invalidNumbers).toBe(2);
	});

	it('writes row records, including rows without cells', () => {
		const sheet = sheetWith([[2, 0, { value: 1 }]]);
		sheet.rows.set(0, { height: 30, customHeight: true });
		sheet.rows.set(2, {
			hidden: true,
			outlineLevel: 2,
			collapsed: true,
			style: 2,
			customFormat: true,
		});
		sheet.rows.set(5, { thickTop: true, thickBot: true, height: 500 });
		sheet.rows.set(6, {});
		const res = buildSheetData(sheet, ctx());
		const rows = pickAll(parseXml(res.xml).root, 'row');
		expect(rows.map((r) => r.attrs)).toEqual([
			{ r: '1', ht: '30', customHeight: '1' },
			{ r: '3', s: '2', customFormat: '1', hidden: '1', outlineLevel: '2', collapsed: '1' },
			{ r: '6', ht: '409', thickTop: '1', thickBot: '1' },
		]);
		expect(res.maxRowOutline).toBe(2);
	});

	it('stores rich text and escapes text', () => {
		const c = ctx();
		const sheet = sheetWith([
			[
				0,
				0,
				{
					value: 'Bold plain',
					richText: [{ text: 'Bold', font: { bold: true } }, { text: ' plain' }],
				},
			],
			[0, 1, { value: ' lead\n_x0041_\u0001' }],
		]);
		buildSheetData(sheet, c);
		const sst = parseXml(c.sst.serialize(true)).root;
		const [rich, plain] = pickAll(sst, 'si');
		expect(childNames(rich)).toEqual(['r', 'r']);
		expect(childNames(pick(rich, 'r', 'rPr'))).toEqual(['b']);
		const t = pick(plain, 't')!;
		expect(t.attrs['xml:space']).toBe('preserve');
		expect(t.children[0]).toBe(' lead\n_x005F_x0041__x0001_');
	});

	it('flags stale formulas', () => {
		expect(
			buildSheetData(sheetWith([[0, 0, { value: 1, formula: 'NOW()', stale: true }]]), ctx()).stale,
		).toBe(true);
	});

	it('serializes 100k cells quickly', () => {
		const sheet = createWorksheet('Big', 1);
		for (let r = 0; r < 2000; r++) {
			for (let col = 0; col < 50; col++) {
				sheet.cells.set(
					cellKey(r, col),
					col % 5 === 0
						? { value: `s${r % 300}` }
						: col % 7 === 0
							? { value: r, formula: `A${r + 1}*2` }
							: { value: r * col + 0.5 },
				);
			}
		}
		const start = performance.now();
		const res = buildSheetData(sheet, ctx());
		const elapsed = performance.now() - start;
		expect(res.used).toEqual({ top: 0, left: 0, bottom: 1999, right: 49 });
		expect(elapsed).toBeLessThan(400);
	});
});

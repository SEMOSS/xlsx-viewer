import { describe, expect, it } from 'vitest';

import { createWorksheet } from '../sheet/worksheet';
import type { Table } from '../types/table';
import { el, parseXml, serializeElement } from '../xml';
import { CellMetadata } from './metadata';
import { SheetGeometry } from './sheet-geometry';
import { buildTableXml, normalizeTableColumns } from './tables';
import { childNames, pick, pickAll } from './test-utils/package-check';
import { ensureNamespace, mergeAttrs, normalizeRoot, setOrderedChild } from './xml-utils';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function table(extra: Partial<Table> = {}): Table {
	return {
		id: 3,
		name: 'Sales',
		displayName: 'Sales',
		ref: { top: 0, left: 0, bottom: 4, right: 2 },
		headerRowCount: 1,
		totalsRowCount: 0,
		columns: [
			{ id: 1, name: 'Region' },
			{ id: 2, name: 'Units', totalsRowFunction: 'sum' },
			{ id: 3, name: 'Price', calculatedColumnFormula: 'XLOOKUP([@Region],A:A,B:B)' },
		],
		...extra,
	};
}

describe('tables', () => {
	it('writes a new table part in CT_Table order', () => {
		const root = parseXml(
			buildTableXml(
				table({
					autoFilter: {
						ref: { top: 0, left: 0, bottom: 4, right: 2 },
						columns: [{ colId: 0, filters: { values: ['East'] } }],
					},
					style: {
						name: 'TableStyleLight9',
						showFirstColumn: false,
						showLastColumn: false,
						showRowStripes: true,
						showColumnStripes: false,
					},
				}),
				7,
			),
		).root;
		expect(root.attrs).toMatchObject({
			xmlns: MAIN,
			id: '7',
			name: 'Sales',
			displayName: 'Sales',
			ref: 'A1:C5',
			totalsRowShown: '0',
		});
		expect(childNames(root)).toEqual(['autoFilter', 'tableColumns', 'tableStyleInfo']);
		const cols = pickAll(pick(root, 'tableColumns'), 'tableColumn');
		expect(cols.map((c) => c.attrs)).toEqual([
			{ id: '1', name: 'Region' },
			{ id: '2', name: 'Units', totalsRowFunction: 'sum' },
			{ id: '3', name: 'Price' },
		]);
		expect(pick(cols[2], 'calculatedColumnFormula')?.children[0]).toBe(
			'_xlfn.XLOOKUP([[#This Row],[Region]],A:A,B:B)',
		);
		expect(pick(root, 'tableStyleInfo')?.attrs).toEqual({
			name: 'TableStyleLight9',
			showFirstColumn: '0',
			showLastColumn: '0',
			showRowStripes: '1',
			showColumnStripes: '0',
		});
	});

	it('regenerates on top of the original markup', () => {
		const raw = parseXml(
			`<table xmlns="${MAIN}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{T}" id="3" name="Old" displayName="Old" ref="A1:C5" totalsRowShown="0"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" xr3:uid="{C1}" name="Region" xmlns:xr3="urn:x"/><tableColumn id="2" name="Units"/><tableColumn id="3" name="Price"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/><extLst><ext uri="{X}"/></extLst></table>`,
		).root;
		const root = parseXml(
			buildTableXml(table({ raw, headerRowCount: 0, totalsRowCount: 1 }), 3),
		).root;
		expect(root.attrs).toMatchObject({
			'xr:uid': '{T}',
			name: 'Sales',
			headerRowCount: '0',
			totalsRowCount: '1',
		});
		// A shown totals row means totalsRowShown takes its default (true).
		expect(root.attrs.totalsRowShown).toBeUndefined();
		expect(childNames(root)).toEqual(['tableColumns', 'extLst']);
		expect(pick(root, 'tableColumns', 'tableColumn')?.attrs['xr3:uid']).toBe('{C1}');
	});

	it('normalises columns to the range width with unique ids and names', () => {
		const cols = normalizeTableColumns(
			table({
				ref: { top: 0, left: 0, bottom: 3, right: 3 },
				columns: [
					{ id: 1, name: 'A' },
					{ id: 1, name: 'a' },
					{ id: 5, name: '' },
				],
			}),
		);
		expect(cols.map((c) => [c.id, c.name])).toEqual([
			[1, 'A'],
			[2, 'a2'],
			[3, 'Column3'],
			[4, 'Column4'],
		]);
	});
});

describe('cell metadata', () => {
	const XLDAPR =
		`<metadata xmlns="${MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
		'<metadataTypes count="2"><metadataType name="XLRICHVALUE"/><metadataType name="XLDAPR"/></metadataTypes>' +
		'<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
		'<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="2" v="0"/></bk></cellMetadata></metadata>';

	it('reuses the dynamic-array record of the source part', () => {
		const md = CellMetadata.parse(XLDAPR, 'xl/metadata.xml');
		expect(md.dynamicArrayCm(undefined)).toBe(2);
		expect(md.dynamicArrayCm(1)).toBe(1);
		expect(md.changed).toBe(false);
	});

	it('adds an XLDAPR record to a source part that lacks one', () => {
		const md = CellMetadata.parse(
			`<metadata xmlns="${MAIN}"><metadataTypes count="1"><metadataType name="XLRICHVALUE"/></metadataTypes><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`,
			'xl/metadata.xml',
		);
		expect(md.dynamicArrayCm(undefined)).toBe(1);
		expect(md.changed).toBe(true);
		const root = parseXml(md.serialize()).root;
		expect(childNames(root)).toEqual([
			'metadataTypes',
			'futureMetadata',
			'cellMetadata',
			'valueMetadata',
		]);
		expect(pick(root, 'metadataTypes')?.attrs.count).toBe('2');
		expect(pick(root, 'cellMetadata', 'bk', 'rc')?.attrs).toEqual({ t: '2', v: '0' });
		expect(md.validVm(1)).toBe(true);
	});
});

describe('sheet geometry', () => {
	it('accounts for custom widths, hidden columns and row heights', () => {
		const s = createWorksheet('S', 1);
		s.cols = [
			{ min: 1, max: 2, width: 20, customWidth: true },
			{ min: 3, max: 3, hidden: true },
		];
		s.rows.set(2, { height: 30 });
		s.rows.set(3, { hidden: true });
		const g = new SheetGeometry(s);
		expect(g.colWidth(0)).toBe(64);
		expect(g.colWidth(1)).toBe(140);
		expect(g.colLeft(3)).toBe(64 + 140 * 2);
		expect(g.colLeft(4)).toBe(64 + 140 * 2);
		expect(g.rowTop(3)).toBe(20 * 2 + 40);
		expect(g.rowTop(5)).toBe(20 * 2 + 40 + 20);
		expect(
			g.anchorBox({
				type: 'oneCell',
				from: { col: 1, colOff: 9525, row: 1, rowOff: 0 },
				ext: { cx: 5, cy: 6 },
			}),
		).toEqual({
			x: 64 * 9525 + 9525,
			y: 20 * 9525,
			cx: 5,
			cy: 6,
		});
	});
});

describe('xml utils', () => {
	it('merges model attributes into original ones in place', () => {
		expect(
			mergeAttrs({ a: '1', zoom: '120', keep: 'x' }, { zoom: 90, a: undefined, added: true }),
		).toEqual({
			zoom: '90',
			keep: 'x',
			added: '1',
		});
	});

	it('inserts in schema order, replacing all same-named children at the first position', () => {
		const order = ['a', 'b', 'c', 'd'];
		const root = el('r', {}, [el('a'), el('c'), el('c'), el('x:unknown'), el('d')]);
		setOrderedChild(root, 'b', el('b'), order);
		setOrderedChild(root, 'c', [el('c', { n: 1 })], order);
		expect(serializeElement(root)).toBe('<r><a/><b/><c n="1"/><x:unknown/><d/></r>');
		setOrderedChild(root, 'c', null, order);
		expect(serializeElement(root)).toBe('<r><a/><b/><x:unknown/><d/></r>');
	});

	it('normalises prefixed roots and declares namespaces once', () => {
		const root = el('x:worksheet', { 'xmlns:x': MAIN });
		expect(normalizeRoot(root, 'worksheet', MAIN)).toBe('x');
		expect(root.name).toBe('worksheet');
		expect(Object.keys(root.attrs)[0]).toBe('xmlns');
		expect(ensureNamespace(root, 'r', 'urn:r')).toBe('r');
		expect(ensureNamespace(root, 'r', 'urn:r')).toBe('r');
		expect(ensureNamespace(root, 'r', 'urn:other')).toBe('r1');
	});
});

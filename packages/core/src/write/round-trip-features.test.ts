import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { readWorkbook } from '../read';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import type { DrawingObject } from '../types/drawing';
import type { Workbook } from '../types/workbook';
import { writeWorkbook } from './index';
import { fixtureBytes } from './test-utils/fixtures';
import {
	childNames,
	openPackage,
	partRoot,
	pick,
	pickAll,
	validatePackage,
} from './test-utils/package-check';

const HR = 'issue-132-hr-deck--Workbook1.xlsx';
const PNG = new Uint8Array(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
		'base64',
	),
);
const ROOT = '{AAAAAAAA-1111-2222-3333-444444444444}';
const BOB = '{11111111-2222-3333-4444-555555555555}';

async function save(wb: Workbook) {
	const bytes = await writeWorkbook(wb, { now: new Date('2025-01-01T00:00:00Z') });
	const pkg = await openPackage(bytes);
	expect(await validatePackage(pkg)).toEqual([]);
	return { bytes, pkg, back: await readWorkbook(bytes) };
}

function twoCell(col: number, row: number, cols = 3, rows = 5) {
	return {
		type: 'twoCell' as const,
		from: { col, colOff: 0, row, rowOff: 0 },
		to: { col: col + cols, colOff: 0, row: row + rows, rowOff: 0 },
	};
}

describe('comments round trip', () => {
	it('adds notes, threaded comments and persons to a loaded sheet', async () => {
		const wb = await readWorkbook(fixtureBytes(HR));
		const s = wb.sheets[0];
		s.notes = [{ ref: { row: 0, col: 0 }, author: 'Ann', text: 'Check\nthis', visible: true }];
		wb.persons = [{ id: BOB, displayName: 'Bob', providerId: 'None' }];
		s.threadedComments = [
			{
				id: ROOT,
				ref: { row: 2, col: 1 },
				personId: BOB,
				text: 'Why?',
				created: '2024-05-01T10:00:00.00',
			},
			{ id: 'reply-1', ref: { row: 2, col: 1 }, personId: BOB, text: 'Because', parentId: ROOT },
		];
		s.commentsDirty = true;
		const { pkg, back } = await save(wb);
		for (const p of [
			'xl/comments1.xml',
			'xl/drawings/vmlDrawing1.vml',
			'xl/threadedComments/threadedComment1.xml',
			'xl/persons/person.xml',
		]) {
			expect(pkg.has(p), p).toBe(true);
		}
		expect(pkg.contentTypes.get('xl/drawings/vmlDrawing1.vml')).toBe(
			'application/vnd.openxmlformats-officedocument.vmlDrawing',
		);
		const sheetXml = await partRoot(pkg, 'xl/worksheets/sheet1.xml');
		expect(childNames(sheetXml).slice(-2)).toEqual(['legacyDrawing', 'tableParts']);
		const comments = await partRoot(pkg, 'xl/comments1.xml');
		expect(pickAll(pick(comments, 'authors'), 'author').map((a) => a.children[0])).toEqual([
			'Ann',
			`tc=${ROOT}`,
		]);
		const mirror = pickAll(pick(comments, 'commentList'), 'comment')[1];
		expect(mirror.attrs).toMatchObject({ ref: 'B3', 'xr:uid': ROOT });
		expect(pick(mirror, 'text', 't')?.children[0]).toContain(
			'Comment:\n    Why?\nReply:\n    Because',
		);
		const vml = (await pkg.text('xl/drawings/vmlDrawing1.vml'))!;
		expect(vml).toContain('ObjectType="Note"');
		expect(vml).toContain('<x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor>');
		expect(vml).toContain('<x:Visible/>');

		const t = back.sheets[0];
		expect(t.notes).toHaveLength(1);
		expect(t.notes[0]).toMatchObject({ author: 'Ann', text: 'Check\nthis', visible: true });
		expect(t.threadedComments).toHaveLength(2);
		const [root, reply] = t.threadedComments;
		expect(root.id).toBe(ROOT);
		expect(reply.parentId).toBe(ROOT);
		expect(reply.id).toMatch(/^\{[0-9A-F-]{36}\}$/);
		expect(back.persons.map((p) => p.displayName)).toEqual(['Bob']);
	});

	it('removes comment parts when the notes are deleted', async () => {
		const wb = createWorkbook();
		wb.sheets[0].notes = [{ ref: { row: 4, col: 4 }, author: 'A', text: 'x' }];
		const first = await save(wb);
		first.back.sheets[0].notes = [];
		first.back.sheets[0].commentsDirty = true;
		const { pkg } = await save(first.back);
		expect(pkg.partPaths().some((p) => p.includes('comments') || p.endsWith('.vml'))).toBe(false);
		expect(pick(await partRoot(pkg, 'xl/worksheets/sheet1.xml'), 'legacyDrawing')).toBeUndefined();
	});
});

describe('drawings round trip', () => {
	it('adds a picture, a shape and a chart, then re-emits untouched objects verbatim', async () => {
		const wb = await readWorkbook(fixtureBytes(HR));
		const s = wb.sheets[0];
		wb.media.set('xl/media/image1.png', {
			path: 'xl/media/image1.png',
			contentType: 'image/png',
			data: PNG,
		});
		s.drawings = [
			{
				kind: 'picture',
				id: '2',
				name: 'Logo',
				mediaPath: 'xl/media/image1.png',
				anchor: {
					type: 'oneCell',
					from: { col: 3, colOff: 0, row: 1, rowOff: 0 },
					ext: { cx: 952500, cy: 952500 },
				},
			},
			{
				kind: 'shape',
				id: '3',
				geometry: 'roundRect',
				anchor: twoCell(5, 1),
				fill: { type: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
				line: { width: 12700, dash: 'dash' },
				text: {
					paragraphs: [
						{ align: 'ctr', runs: [{ text: 'Hello\nworld', style: { bold: true, size: 14 } }] },
					],
				},
			},
			{
				kind: 'chart',
				id: '4',
				name: 'Chart 1',
				anchor: twoCell(0, 8, 6, 12),
				chart: {
					plots: [
						{
							type: 'bar',
							barDirection: 'col',
							grouping: 'clustered',
							series: [
								{
									idx: 0,
									order: 0,
									values: { kind: 'num', ref: 'Sheet1!$B$2:$B$6' },
									categories: { kind: 'str', ref: 'Sheet1!$A$2:$A$6' },
								},
							],
							axisIds: [1, 2],
						},
					],
					axes: [
						{ id: 1, kind: 'cat', position: 'b', crossAxisId: 2 },
						{ id: 2, kind: 'val', position: 'l', crossAxisId: 1 },
					],
				},
			},
		];
		s.drawingsDirty = true;
		const first = await save(wb);
		const pkg = first.pkg;
		expect(pkg.has('xl/drawings/drawing1.xml')).toBe(true);
		expect(await pkg.bytes('xl/media/image1.png')).toEqual(PNG);
		expect(pkg.contentTypes.defaults.get('png')).toBe('image/png');
		expect(pick(await partRoot(pkg, 'xl/worksheets/sheet1.xml'), 'drawing')).toBeDefined();
		const drawing = await partRoot(pkg, 'xl/drawings/drawing1.xml');
		expect(childNames(drawing)).toEqual(['oneCellAnchor', 'twoCellAnchor', 'twoCellAnchor']);
		const ids = (await pkg.text('xl/drawings/drawing1.xml'))!.match(/cNvPr id="(\d+)"/g)!;
		expect(new Set(ids).size).toBe(ids.length);
		expect(pkg.partPaths().filter((p) => p.startsWith('xl/charts/'))).toEqual([
			'xl/charts/chart1.xml',
		]);
		expect(pkg.contentTypes.get('xl/charts/chart1.xml')).toBe(
			'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
		);
		const kinds = first.back.sheets[0].drawings.map((d) => d.kind);
		expect(kinds).toEqual(['picture', 'shape', 'chart']);
		const chart = first.back.sheets[0].drawings[2];
		expect(chart.kind === 'chart' && chart.chart.plots[0].type).toBe('bar');
		const shape = first.back.sheets[0].drawings[1];
		expect(shape.kind === 'shape' && shape.geometry).toBe('roundRect');
		expect(
			shape.kind === 'shape' && shape.text?.paragraphs[0].runs.map((r) => r.text).join('|'),
		).toContain('Hello');

		// Second save: only a new shape is added; the loaded objects keep their markup.
		const wb2 = first.back;
		const loaded = wb2.sheets[0].drawings;
		expect(loaded.every((o) => o.dirty === false && o.raw)).toBe(true);
		const extra: DrawingObject = {
			kind: 'shape',
			id: 'new',
			geometry: 'ellipse',
			anchor: twoCell(9, 1, 2, 2),
		};
		wb2.sheets[0].drawings = [...loaded, extra];
		wb2.sheets[0].drawingsDirty = true;
		const second = await save(wb2);
		const before = (await pkg.text('xl/drawings/drawing1.xml'))!;
		const after = (await second.pkg.text('xl/drawings/drawing1.xml'))!;
		const firstAnchor = before.slice(
			before.indexOf('<xdr:oneCellAnchor'),
			before.indexOf('</xdr:oneCellAnchor>'),
		);
		expect(after).toContain(firstAnchor);
		expect(second.back.sheets[0].drawings.map((d) => d.kind)).toEqual([...kinds, 'shape']);
		const pic = second.back.sheets[0].drawings[0];
		expect(pic.kind === 'picture' && second.back.media.get(pic.mediaPath)?.data).toEqual(PNG);
	});

	it('drops the drawing part and its media when all objects are removed', async () => {
		const wb = createWorkbook();
		wb.media.set('xl/media/image1.png', {
			path: 'xl/media/image1.png',
			contentType: 'image/png',
			data: PNG,
		});
		wb.sheets[0].drawings = [
			{ kind: 'picture', id: '1', mediaPath: 'xl/media/image1.png', anchor: twoCell(0, 0) },
		];
		const first = await save(wb);
		first.back.sheets[0].drawings = [];
		first.back.sheets[0].drawingsDirty = true;
		const { pkg } = await save(first.back);
		expect(pkg.partPaths().filter((p) => p.includes('drawing') || p.includes('media'))).toEqual([]);
	});
});

describe('workbook features round trip', () => {
	it('round-trips tables, conditional formats with inline formats, names and arrays', async () => {
		const wb = await readWorkbook(fixtureBytes(HR));
		const s = wb.sheets[0];
		const dxfCount = wb.styles.dxfs.length;
		s.tables.push({
			id: 1,
			name: 'Extra',
			displayName: 'Extra',
			ref: { top: 9, left: 0, bottom: 11, right: 1 },
			headerRowCount: 1,
			totalsRowCount: 0,
			columns: [
				{ id: 1, name: 'K' },
				{ id: 2, name: 'V', calculatedColumnFormula: 'IFNA(A11,0)' },
			],
			autoFilter: { ref: { top: 9, left: 0, bottom: 11, right: 1 }, columns: [] },
			style: {
				name: 'TableStyleMedium9',
				showFirstColumn: false,
				showLastColumn: false,
				showRowStripes: true,
				showColumnStripes: false,
			},
		});
		putCell(s, 9, 0, { value: 'K' });
		putCell(s, 9, 1, { value: 'V' });
		s.conditionalFormats = [
			{
				ranges: [{ top: 1, left: 1, bottom: 5, right: 1 }],
				rules: [
					{
						type: 'cellIs',
						priority: 1,
						operator: 'greaterThan',
						formulas: ['0.2'],
						dxf: { font: { bold: true, color: { rgb: 'FF9C0006' } } },
					},
				],
			},
		];
		putCell(s, 12, 0, {
			value: 1,
			formula: 'SEQUENCE(3)',
			array: { ref: { top: 12, left: 0, bottom: 14, right: 0 }, dynamic: true },
		});
		putCell(s, 13, 0, { value: 2, arrayAnchor: cellKey(12, 0) });
		putCell(s, 14, 0, { value: 3, arrayAnchor: cellKey(12, 0) });
		putCell(s, 15, 0, { value: 9, formula: 'NOW()', stale: true });
		wb.definedNames.push(
			{ name: 'Total', formula: 'SUM(Sheet1!$B$2:$B$6)+IFNA(1,0)' },
			{ name: 'Local', formula: 'Sheet1!$A$1', localSheet: 0, hidden: true },
		);
		wb.dirty = true;
		const { pkg, back } = await save(wb);

		expect(wb.styles.dxfs).toHaveLength(dxfCount);
		const tables = pkg.partPaths().filter((p) => p.startsWith('xl/tables/'));
		expect(tables).toHaveLength(2);
		const t = back.sheets[0].tables.find((x) => x.name === 'Extra')!;
		expect(t.id).toBe(2);
		expect(t.columns[1].calculatedColumnFormula).toBe('IFNA(A11,0)');
		expect(await pkg.text(tables.find((p) => p !== 'xl/tables/table1.xml')!)).toContain(
			'_xlfn.IFNA(A11,0)',
		);

		const rule = back.sheets[0].conditionalFormats[0].rules[0];
		expect(rule.dxfId).toBe(dxfCount);
		expect(back.styles.dxfs[rule.dxfId!].font).toMatchObject({ bold: true });

		const anchor = back.sheets[0].cells.get(cellKey(12, 0))!;
		expect(anchor.array).toEqual({
			ref: { top: 12, left: 0, bottom: 14, right: 0 },
			dynamic: true,
		});
		expect(pkg.has('xl/metadata.xml')).toBe(true);
		expect(pkg.contentTypes.get('xl/metadata.xml')).toContain('sheetMetadata');

		const wbXml = await partRoot(pkg, 'xl/workbook.xml');
		expect(pick(wbXml, 'calcPr')?.attrs.fullCalcOnLoad).toBe('1');
		const names = pickAll(pick(wbXml, 'definedNames'), 'definedName');
		expect(
			names.map((n) => [n.attrs.name, n.attrs.localSheetId, n.attrs.hidden, n.children[0]]),
		).toEqual([
			['Total', undefined, undefined, 'SUM(Sheet1!$B$2:$B$6)+_xlfn.IFNA(1,0)'],
			['Local', '0', '1', 'Sheet1!$A$1'],
		]);
		expect(back.definedNames.find((n) => n.name === 'Total')?.formula).toBe(
			'SUM(Sheet1!$B$2:$B$6)+IFNA(1,0)',
		);
	});

	it('round-trips sheet features of an in-memory workbook', async () => {
		const wb = createWorkbook({ sheetNames: ['Main', 'Target'] });
		wb.date1904 = true;
		wb.protection = { lockStructure: true };
		const s = wb.sheets[0];
		putCell(s, 0, 0, { value: 'Link' });
		putCell(s, 1, 0, { value: 'Jump' });
		s.hyperlinks = [
			{ ref: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/?q=1&r=2' },
			{ ref: { top: 1, left: 0, bottom: 1, right: 0 }, location: 'Target!B2' },
		];
		s.merges = [{ top: 3, left: 0, bottom: 4, right: 2 }];
		s.view = {
			...s.view,
			pane: { xSplit: 0, ySplit: 1, state: 'frozen' },
			selection: { activeCell: { row: 5, col: 1 }, ranges: [] },
		};
		s.protection = { sheet: true, objects: true };
		s.dataValidations = [
			{
				ranges: [{ top: 2, left: 2, bottom: 9, right: 2 }],
				type: 'list',
				formula1: '"Yes,No"',
				inCellDropdown: true,
			},
		];
		s.autoFilter = { ref: { top: 0, left: 0, bottom: 9, right: 2 }, columns: [] };
		s.cols = [{ min: 0, max: 1, width: 20, customWidth: true }];
		s.rows.set(0, { height: 30, customHeight: true });
		const { back, pkg } = await save(wb);
		expect(back.date1904).toBe(true);
		expect(back.protection?.lockStructure).toBe(true);
		const t = back.sheets[0];
		expect(t.hyperlinks.map((h) => h.target ?? h.location)).toEqual([
			'https://example.com/?q=1&r=2',
			'Target!B2',
		]);
		expect(t.merges).toEqual([{ top: 3, left: 0, bottom: 4, right: 2 }]);
		expect(t.view.pane).toMatchObject({ ySplit: 1, state: 'frozen', activePane: 'bottomLeft' });
		expect(t.view.tabSelected).toBe(true);
		expect(t.protection?.sheet).toBe(true);
		expect(t.dataValidations[0]).toMatchObject({ type: 'list', formula1: '"Yes,No"' });
		expect(t.autoFilter?.ref).toEqual({ top: 0, left: 0, bottom: 9, right: 2 });
		expect(t.cols[0]).toMatchObject({ min: 0, max: 1, width: 20 });
		expect(t.rows.get(0)).toMatchObject({ height: 30, customHeight: true });
		const names = pickAll(
			pick(await partRoot(pkg, 'xl/workbook.xml'), 'definedNames'),
			'definedName',
		);
		expect(names[0].attrs).toMatchObject({
			name: '_xlnm._FilterDatabase',
			localSheetId: '0',
			hidden: '1',
		});
		expect(names[0].children[0]).toBe('Main!$A$1:$C$10');
	});
});

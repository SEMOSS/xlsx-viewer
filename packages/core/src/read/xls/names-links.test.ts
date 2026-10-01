import { describe, expect, it } from 'vitest';

import { readXlsWorkbook } from './index';
import { fileTarget } from './sheet-links';
import {
	buildXls,
	definedName,
	externSheet,
	font,
	supbookExternal,
	supbookSelf,
} from './test-utils/biff-builder';
import { ptg } from './test-utils/ptg-builder';
import { hlink, hlinkTooltip, note, obj, txo } from './test-utils/sheet-builder';
import { RT } from './biff-opcodes';
import { record } from './test-utils/bytes';

const sheets = [
	{ name: 'Data', records: [] },
	{ name: 'Chart', records: [], type: 2, bofType: 0x20 },
	{ name: 'Report Q1', records: [] },
];
const abs = { rowAbs: true, colAbs: true };

describe('defined names', () => {
	it('reads global, sheet-scoped, built-in and hidden names', () => {
		const wb = readXlsWorkbook(
			buildXls(
				[
					supbookSelf(3),
					supbookExternal('Book2.xls', ['S']),
					externSheet([
						[0, 0, 0],
						[0, 2, 2],
						[1, 0, 0],
						[0, 0, 2],
					]),
					definedName({ name: 'Rate', rgce: ptg.ref3d(0, 0, 1, abs) }),
					definedName({ name: 6, itab: 3, rgce: ptg.area3d(1, 0, 0, 9, 3, abs) }),
					definedName({
						name: 7,
						itab: 1,
						rgce: [
							...ptg.area3d(0, 0, 0, 65535, 1, abs),
							...ptg.area3d(0, 0, 0, 1, 255, abs),
							...ptg.union,
						],
					}),
					definedName({ name: 13, itab: 1, hidden: true, rgce: ptg.area3d(0, 0, 0, 20, 2, abs) }),
					definedName({
						name: 'Local',
						itab: 3,
						rgce: [...ptg.ref3d(1, 0, 0, abs), ...ptg.int(2), ...ptg.mul],
					}),
					definedName({
						name: 'Total',
						rgce: [...ptg.area3d(3, 0, 0, 0, 0, abs), ...ptg.funcVar(4, 1)],
					}),
					definedName({ name: 'Next', rgce: ptg.ref3d(0, 1, 0, { colAbs: true }) }),
					definedName({ name: 'External', rgce: ptg.ref3d(2, 0, 0) }),
					definedName({ name: 'OnChart', itab: 2, rgce: ptg.int(1) }),
					definedName({ name: 'MyMacro', func: true, rgce: [] }),
					definedName({ name: '_xlfn.IFERROR', rgce: ptg.err(0x1d), hidden: true }),
					definedName({ name: 'Auto', rgce: [...ptg.name(1), ...ptg.int(1), ...ptg.add] }),
					definedName({ name: 2, rgce: ptg.ref3d(0, 0, 0) }),
				],
				sheets,
			),
		);
		expect(wb.definedNames).toEqual([
			{ name: 'Rate', formula: 'Data!$B$1' },
			{ name: '_xlnm.Print_Area', formula: "'Report Q1'!$A$1:$D$10", localSheet: 1, builtin: true },
			{
				name: '_xlnm.Print_Titles',
				formula: 'Data!$A:$B,Data!$1:$2',
				localSheet: 0,
				builtin: true,
			},
			{
				name: '_xlnm._FilterDatabase',
				formula: 'Data!$A$1:$C$21',
				localSheet: 0,
				hidden: true,
				builtin: true,
			},
			{ name: 'Local', formula: "'Report Q1'!$A$1*2", localSheet: 1 },
			{ name: 'Total', formula: "SUM('Data:Report Q1'!$A$1:$A$1)" },
			{ name: 'Next', formula: 'Data!$A2' },
			{ name: 'Auto', formula: 'Rate+1' },
			{ name: 'Auto_Close', formula: 'Data!A1' },
		]);
		const codes = wb.warnings.map((w) => w.code);
		expect(codes).toContain('external-link');
		expect(codes).toContain('name-skipped');
	});
});

describe('hyperlinks', () => {
	it('reads URL, file and internal links with display text and tooltips', () => {
		const wb = readXlsWorkbook(
			buildXls(
				[],
				[
					{
						name: 'Links',
						records: [
							hlink([0, 0, 0, 0], { url: 'https://example.com/a?b=1', display: 'Example' }),
							hlinkTooltip([0, 0, 0, 0], 'Go to example'),
							hlink([1, 1, 0, 1], { location: 'Links!C5' }),
							hlink([2, 2, 0, 0], { file: { path: 'docs\\report.xls', up: 2 } }),
							hlink([3, 3, 0, 0], { url: 'mailto:someone@example.com', location: 'ignored' }),
							hlink([4, 4, 0, 0], {}),
						],
					},
				],
			),
		);
		expect(wb.sheets[0].hyperlinks).toEqual([
			{
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'https://example.com/a?b=1',
				display: 'Example',
				tooltip: 'Go to example',
			},
			{ ref: { top: 1, left: 0, bottom: 1, right: 1 }, location: 'Links!C5' },
			{ ref: { top: 2, left: 0, bottom: 2, right: 0 }, target: '../../docs/report.xls' },
			{
				ref: { top: 3, left: 0, bottom: 3, right: 0 },
				target: 'mailto:someone@example.com',
				location: 'ignored',
			},
		]);
	});
});

describe('file link targets', () => {
	it('turns absolute and UNC paths into file URLs and relative paths into slashes', () => {
		expect(fileTarget('C:\\Docs\\a.xls')).toBe('file:///C:\\Docs\\a.xls');
		expect(fileTarget('\\\\srv\\share\\a.xls')).toBe('file:///\\\\srv\\share\\a.xls');
		expect(fileTarget('..\\sub\\a.xls')).toBe('../sub/a.xls');
	});
});

describe('notes', () => {
	it('reads note text, author, visibility and formatting runs', () => {
		const wb = readXlsWorkbook(
			buildXls(
				[font({ name: 'Tahoma', height: 160, bold: true }), font({ name: 'Tahoma', height: 160 })],
				[
					{
						name: 'Notes',
						records: [
							obj(0x19, 1),
							txo('Ann:\nCheck this', [
								[0, 5],
								[4, 6],
							]),
							obj(0x19, 2),
							txo('Ünïcødé note €', [[0, 6]], true),
							obj(0x19, 3),
							record(RT.TXO, new Array<number>(18).fill(0)),
							note(2, 1, 1, 'Ann', true),
							note(0, 0, 2, 'Bob'),
							note(5, 5, 3, ''),
							note(6, 6, 99, 'Nobody'),
						],
					},
				],
			),
		);
		const notes = wb.sheets[0].notes;
		expect(notes).toHaveLength(4);
		expect(notes[0]).toEqual({
			ref: { row: 2, col: 1 },
			author: 'Ann',
			text: 'Ann:\nCheck this',
			runs: [
				{ text: 'Ann:', font: { name: 'Tahoma', size: 8, bold: true, family: 2 } },
				{ text: '\nCheck this', font: { name: 'Tahoma', size: 8, family: 2 } },
			],
			visible: true,
		});
		expect(notes[1]).toEqual({
			ref: { row: 0, col: 0 },
			author: 'Bob',
			text: 'Ünïcødé note €',
		});
		expect(notes[2]).toEqual({ ref: { row: 5, col: 5 }, author: '', text: '' });
		expect(notes[3].text).toBe('');
	});
});

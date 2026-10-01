import { describe, expect, it } from 'vitest';

import { cellKey } from '../../address/cell-address';
import type { Cell } from '../../types/cell';
import type { Workbook } from '../../types/workbook';
import { writeWorkbook } from '../../write';
import { readWorkbook } from '../index';
import { readXlsWorkbook } from './index';
import {
	buildXls,
	definedName,
	externSheet,
	font,
	format,
	palette,
	sst,
	supbookAddin,
	externName,
	supbookSelf,
	window1,
	xf,
} from './test-utils/biff-builder';
import { arrayExtra, ptg, rgce } from './test-utils/ptg-builder';
import {
	arrayFormula,
	boolErr,
	cf,
	colInfo,
	condFmt,
	dv,
	formula,
	hlink,
	labelSst,
	mergedCells,
	note,
	number,
	obj,
	pane,
	rk,
	rkInt,
	row,
	shrfmla,
	stringResult,
	txo,
	window2,
} from './test-utils/sheet-builder';

const abs = { rowAbs: true, colAbs: true };

function sampleXls(): Uint8Array {
	const globals = [
		font({ name: 'Calibri', height: 220, bold: true, color: 10 }),
		format(164, '#,##0.000'),
		xf({
			font: 5,
			numFmt: 164,
			fill: { fls: 1, fore: 13, back: 64 },
			borders: [
				[1, 8],
				[1, 8],
				[2, 8],
				[6, 12],
			],
			h: 2,
			wrap: true,
		}),
		xf({ numFmt: 14 }),
		palette(['336699']),
		sst(['Name', 'Amount', 'Apples', 'Pears', 'Total']),
		supbookSelf(3),
		supbookAddin(),
		externName('EDATE'),
		externSheet([
			[0, 0, 0],
			[0, 1, 1],
			[1, 0xfffe, 0xfffe],
		]),
		definedName({ name: 'Amounts', rgce: ptg.area3d(0, 1, 1, 2, 1, abs) }),
		definedName({ name: 6, itab: 1, rgce: ptg.area3d(0, 0, 0, 5, 3, abs) }),
		window1({ activeTab: 1 }),
	];
	const data = [
		colInfo(0, 0, 4000, { custom: true }),
		row(0, { height: 400, custom: true }),
		labelSst(0, 0, 16, 0),
		labelSst(0, 1, 16, 1),
		labelSst(1, 0, 15, 2),
		rk(1, 1, 15, rkInt(12)),
		labelSst(2, 0, 15, 3),
		number(2, 1, 15, 7.5),
		labelSst(3, 0, 15, 4),
		formula(3, 1, 15, 19.5, rgce(ptg.area(1, 1, 2, 1), ptg.attrSum)),
		// Shared formula C2:C3 = B2*2.
		formula(1, 2, 15, 24, rgce(ptg.exp(1, 2))),
		shrfmla(1, 2, 2, 2, rgce(ptg.refN(0, -1), ptg.int(2), ptg.mul)),
		formula(2, 2, 15, 15, rgce(ptg.exp(1, 2))),
		// Array formula D2:D3 = B2:B3*10.
		formula(1, 3, 15, 120, rgce(ptg.exp(1, 3))),
		arrayFormula(1, 2, 3, 3, rgce(ptg.area(1, 1, 2, 1, { cls: 0x60 }), ptg.int(10), ptg.mul)),
		formula(2, 3, 15, 75, rgce(ptg.exp(1, 3))),
		formula(
			4,
			0,
			17,
			39448,
			rgce(ptg.nameX(2, 1), ptg.int(39417), ptg.int(1), ptg.funcVar(255, 3)),
		),
		formula(4, 1, 15, 'Apples!', rgce(ptg.ref(1, 0), ptg.str('!'), ptg.concat)),
		stringResult('Apples!'),
		formula(4, 2, 15, true, rgce(ptg.ref(1, 1), ptg.int(10), ptg.gt)),
		formula(4, 3, 15, 2, rgce(ptg.array(), ptg.funcVar(0, 1)), arrayExtra([[1, 'x', 3]])),
		boolErr(5, 0, 15, 0x07, true),
		mergedCells([[6, 7, 0, 1]]),
		window2({ flags: 0x06be | 0x0100 }),
		pane(0, 1, 1, 0, 2),
		hlink([0, 0, 0, 0], { url: 'https://example.com/' }),
		hlink([1, 1, 0, 0], { location: 'Summary!A1' }),
		obj(0x19, 1),
		txo('Check this'),
		note(1, 1, 1, 'Reviewer'),
		dv({ type: 3, list: true, f1: ptg.str('Apples\u0000Pears'), ranges: [[1, 2, 0, 0]] }),
		condFmt(1, [[1, 2, 1, 1]]),
		cf({ kind: 1, op: 5, f1: ptg.int(10), fill: 10 }),
	];
	const summary = [
		formula(0, 0, 15, 19.5, rgce(ptg.ref3d(0, 3, 1, abs))),
		formula(1, 0, 15, 19.5, rgce(ptg.name(1), ptg.funcVar(4, 1))),
		formula(2, 0, 15, 2, rgce(ptg.area3d(0, 1, 1, 2, 1), ptg.funcVar(0, 1))),
	];
	return buildXls(globals, [
		{ name: 'Data', records: data },
		{ name: 'Summary', records: summary },
		{ name: 'Archive', records: [number(0, 0, 15, 42)], state: 'hidden' },
	]);
}

function comparable(cell: Cell | undefined): Partial<Cell> | undefined {
	if (!cell) {
		return undefined;
	}
	const out: Partial<Cell> = { value: cell.value };
	if (cell.formula !== undefined) {
		out.formula = cell.formula;
	}
	if (cell.array) {
		out.array = cell.array;
	}
	return out;
}

function cellsOf(wb: Workbook, sheet: number): Map<number, Partial<Cell> | undefined> {
	const out = new Map<number, Partial<Cell> | undefined>();
	for (const [key, cell] of wb.sheets[sheet].cells) {
		if (cell.value !== null || cell.formula !== undefined) {
			out.set(key, comparable(cell));
		}
	}
	return out;
}

describe('.xls -> .xlsx round trip', () => {
	it('writes an .xlsx that reads back with the same content', async () => {
		const xls = readXlsWorkbook(sampleXls());
		expect(xls.warnings).toEqual([]);
		const back = await readWorkbook(await writeWorkbook(xls));
		expect(back.warnings).toEqual([]);

		expect(back.sheets.map((s) => [s.name, s.state])).toEqual([
			['Data', 'visible'],
			['Summary', 'visible'],
			['Archive', 'hidden'],
		]);
		for (const i of [0, 1, 2]) {
			expect(cellsOf(back, i)).toEqual(cellsOf(xls, i));
		}
		const data = back.sheets[0];
		const at = (r: number, c: number): Cell | undefined => data.cells.get(cellKey(r, c));
		expect(at(3, 1)?.formula).toBe('SUM(B2:B3)');
		expect(at(2, 2)?.formula).toBe('B3*2');
		expect(at(1, 3)?.array).toEqual({
			ref: { top: 1, left: 3, bottom: 2, right: 3 },
			dynamic: false,
		});
		expect(at(4, 0)?.formula).toBe('EDATE(39417,1)');
		expect(at(4, 3)?.formula).toBe('COUNT({1,"x",3})');
		expect(back.sheets[1].cells.get(cellKey(0, 0))?.formula).toBe('Data!$B$4');
		expect(back.sheets[1].cells.get(cellKey(1, 0))?.formula).toBe('SUM(Amounts)');

		// Formatting.
		const header = at(0, 0)!;
		const xfEntry = back.styles.cellXfs[header.style!];
		expect(back.styles.fonts[xfEntry.fontId]).toMatchObject({
			name: 'Calibri',
			size: 11,
			bold: true,
			color: { indexed: 10 },
		});
		expect(back.styles.numFmts[xfEntry.numFmtId]).toBe('#,##0.000');
		expect(back.styles.fills[xfEntry.fillId]).toMatchObject({
			patternType: 'solid',
			fgColor: { indexed: 13 },
		});
		expect(back.styles.borders[xfEntry.borderId].bottom).toEqual({
			style: 'double',
			color: { indexed: 12 },
		});
		expect(xfEntry.alignment).toEqual({ horizontal: 'center', wrapText: true });
		expect(back.styles.cellXfs[at(4, 0)!.style!].numFmtId).toBe(14);
		expect(back.styles.indexedColors?.[8]).toBe('FF336699');

		// Layout and view.
		expect(data.merges).toEqual([{ top: 6, left: 0, bottom: 7, right: 1 }]);
		expect(data.cols[0]).toMatchObject({ min: 0, max: 0, width: 4000 / 256, customWidth: true });
		expect(data.rows.get(0)).toMatchObject({ height: 20, customHeight: true });
		expect(data.view.pane).toMatchObject({
			ySplit: 1,
			state: 'frozen',
			topLeftCell: { row: 1, col: 0 },
		});
		expect(back.view.activeTab).toBe(1);

		// Names, links, notes and rules.
		expect(back.definedNames).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ name: 'Amounts', formula: 'Data!$B$2:$B$3' }),
				expect.objectContaining({
					name: '_xlnm.Print_Area',
					formula: 'Data!$A$1:$D$6',
					localSheet: 0,
				}),
			]),
		);
		expect(data.hyperlinks).toEqual([
			expect.objectContaining({
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'https://example.com/',
			}),
			expect.objectContaining({
				ref: { top: 1, left: 0, bottom: 1, right: 0 },
				location: 'Summary!A1',
			}),
		]);
		expect(data.notes).toEqual([
			expect.objectContaining({ ref: { row: 1, col: 1 }, author: 'Reviewer', text: 'Check this' }),
		]);
		expect(data.dataValidations).toEqual([
			expect.objectContaining({
				type: 'list',
				formula1: '"Apples,Pears"',
				ranges: [{ top: 1, left: 0, bottom: 2, right: 0 }],
			}),
		]);
		expect(data.conditionalFormats).toEqual([
			expect.objectContaining({
				ranges: [{ top: 1, left: 1, bottom: 2, right: 1 }],
				rules: [
					expect.objectContaining({ type: 'cellIs', operator: 'greaterThan', formulas: ['10'] }),
				],
			}),
		]);
		const rule = data.conditionalFormats[0].rules[0];
		expect(back.styles.dxfs[rule.dxfId!].fill).toMatchObject({ bgColor: { indexed: 10 } });
	});
});

describe('readWorkbook entry point', () => {
	it('opens .xls files through readWorkbook', async () => {
		const viaEntry = await readWorkbook(sampleXls());
		const direct = readXlsWorkbook(sampleXls());
		expect(viaEntry.sheets.map((s) => s.name)).toEqual(['Data', 'Summary', 'Archive']);
		for (const i of [0, 1, 2]) {
			expect(cellsOf(viaEntry, i)).toEqual(cellsOf(direct, i));
		}
	});
});

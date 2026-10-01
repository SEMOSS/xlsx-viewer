import { describe, expect, it } from 'vitest';

import { cellKey } from '../../address/cell-address';
import type { Cell } from '../../types/cell';
import type { Workbook } from '../../types/workbook';
import { ERRORS } from '../../values/errors';
import { readXlsWorkbook } from './index';
import { decodeRk } from './sheet-cells';
import {
	buildXls,
	continued,
	externName,
	externSheet,
	font,
	sst,
	supbookExternal,
	supbookSelf,
	xf,
} from './test-utils/biff-builder';
import { u16, u32, xlString } from './test-utils/bytes';
import { ptg, rgce } from './test-utils/ptg-builder';
import {
	arrayFormula,
	blank,
	boolErr,
	formula,
	label,
	labelSst,
	mulblank,
	mulrk,
	number,
	rk,
	rkFloat,
	rkInt,
	rstring,
	shrfmla,
	stringResult,
	table,
} from './test-utils/sheet-builder';
import { RT } from './biff-opcodes';

function cellsOf(wb: Workbook, sheet = 0): Map<number, Cell> {
	return wb.sheets[sheet].cells;
}

function at(wb: Workbook, row: number, col: number, sheet = 0): Cell | undefined {
	return cellsOf(wb, sheet).get(cellKey(row, col));
}

function sheetWith(records: number[][], globals: number[][] = []): Workbook {
	return readXlsWorkbook(buildXls(globals, [{ name: 'Sheet1', records }]));
}

describe('RK numbers', () => {
	it.each<[string, number, number]>([
		['integer', rkInt(42), 42],
		['negative integer', rkInt(-7), -7],
		['largest 30-bit integer', rkInt(0x1fffffff), 0x1fffffff],
		['integer / 100', rkInt(12345, true), 123.45],
		['IEEE double', rkFloat(1.5), 1.5],
		['negative IEEE double', rkFloat(-2.5), -2.5],
		['IEEE double / 100', rkFloat(1, true), 0.01],
		['large IEEE double', rkFloat(2 ** 100), 2 ** 100],
	])('%s', (_label, bits, expected) => {
		expect(decodeRk(bits)).toBe(expected);
	});
});

describe('value cells', () => {
	it('reads every value record type', () => {
		const wb = sheetWith(
			[
				number(0, 0, 15, 3.25),
				rk(0, 1, 15, rkInt(10)),
				mulrk(1, 0, [
					[15, rkInt(1)],
					[15, rkFloat(2.5)],
					[15, rkInt(300, true)],
				]),
				labelSst(2, 0, 15, 1),
				label(2, 1, 15, 'Label €'),
				rstring(2, 2, 15, 'Rich text', [
					[0, 1],
					[5, 0],
				]),
				boolErr(3, 0, 15, 1, false),
				boolErr(3, 1, 15, 0x07, true),
				boolErr(3, 2, 15, 0x2a, true),
				blank(4, 0, 16),
				blank(4, 1, 15),
				mulblank(5, 2, [16, 15, 16]),
			],
			[sst(['zero', 'one']), xf({ numFmt: 10 })],
		);
		expect(at(wb, 0, 0)).toEqual({ value: 3.25 });
		expect(at(wb, 0, 1)).toEqual({ value: 10 });
		expect([0, 1, 2].map((c) => at(wb, 1, c)?.value)).toEqual([1, 2.5, 3]);
		expect(at(wb, 2, 0)).toEqual({ value: 'one' });
		expect(at(wb, 2, 1)).toEqual({ value: 'Label €' });
		expect(at(wb, 2, 2)?.value).toBe('Rich text');
		expect(at(wb, 2, 2)?.richText).toEqual([
			{ text: 'Rich ', font: { name: 'Arial', size: 10, bold: true, family: 2 } },
			{ text: 'text', font: { name: 'Arial', size: 10, family: 2 } },
		]);
		expect(at(wb, 3, 0)).toEqual({ value: true });
		expect(at(wb, 3, 1)?.value).toBe(ERRORS.DIV0);
		expect(at(wb, 3, 2)?.value).toBe(ERRORS.NA);
		// Blank cells are kept only when they carry a format.
		expect(at(wb, 4, 0)).toEqual({ value: null, style: 1 });
		expect(at(wb, 4, 1)).toBeUndefined();
		expect(at(wb, 5, 2)).toEqual({ value: null, style: 1 });
		expect(at(wb, 5, 3)).toBeUndefined();
		expect(at(wb, 5, 4)).toEqual({ value: null, style: 1 });
	});

	it('reads rich shared strings', () => {
		const sstRecord = continued(RT.SST, [
			[
				...u32(1),
				...u32(1),
				...u16(5),
				0x08,
				...u16(1),
				...[...'AbCdE'].map((c) => c.charCodeAt(0)),
				...u16(2),
				...u16(1),
			],
		]);
		const wb = sheetWith([labelSst(0, 0, 15, 0)], [sstRecord]);
		expect(at(wb, 0, 0)?.richText).toEqual([
			{ text: 'Ab' },
			{ text: 'CdE', font: { name: 'Arial', size: 10, bold: true, family: 2 } },
		]);
	});

	it('reports missing shared strings and bad formats without failing', () => {
		const wb = sheetWith([labelSst(0, 0, 15, 5), number(0, 1, 999, 1)], [sst(['a'])]);
		expect(at(wb, 0, 0)).toEqual({ value: '' });
		expect(at(wb, 0, 1)).toEqual({ value: 1 });
		expect(wb.warnings.map((w) => w.code)).toEqual(
			expect.arrayContaining(['missing-shared-string', 'missing-style']),
		);
	});

	it('maps NaN and infinite numbers to #NUM!', () => {
		const wb = sheetWith([number(0, 0, 15, Number.POSITIVE_INFINITY)]);
		expect(at(wb, 0, 0)?.value).toBe(ERRORS.NUM);
	});
});

describe('formula cells', () => {
	it('keeps every kind of cached result', () => {
		const one = rgce(ptg.int(1));
		const wb = sheetWith([
			formula(0, 0, 15, 2.5, rgce(ptg.num(1.25), ptg.int(2), ptg.mul)),
			formula(1, 0, 15, 'x', rgce(ptg.str('x'))),
			stringResult('x'),
			formula(2, 0, 15, true, rgce(ptg.bool(true))),
			formula(3, 0, 15, { error: 0x07 }, rgce(ptg.int(1), ptg.int(0), ptg.div)),
			formula(4, 0, 15, '', rgce(ptg.str(''))),
			formula(5, 0, 15, 'long', one),
			// A record between FORMULA and STRING breaks the pairing: the stray STRING is ignored.
			number(6, 0, 15, 1),
			stringResult('stray'),
		]);
		expect(at(wb, 0, 0)).toEqual({ value: 2.5, formula: '1.25*2' });
		expect(at(wb, 1, 0)).toEqual({ value: 'x', formula: '"x"' });
		expect(at(wb, 2, 0)).toEqual({ value: true, formula: 'TRUE' });
		expect(at(wb, 3, 0)).toEqual({ value: ERRORS.DIV0, formula: '1/0' });
		expect(at(wb, 4, 0)).toEqual({ value: '', formula: '""' });
		expect(at(wb, 5, 0)).toEqual({ value: '', formula: '1' });
		expect(at(wb, 6, 0)).toEqual({ value: 1 });
	});

	it('expands shared formulas into per-cell formulas', () => {
		const shared = rgce(
			ptg.refN(0, -1),
			ptg.refN(0, -2),
			ptg.add,
			ptg.ref(0, 0, { rowAbs: true, colAbs: true }),
			ptg.mul,
		);
		const records = [
			formula(1, 2, 15, 3, rgce(ptg.exp(1, 2))),
			shrfmla(1, 3, 2, 2, shared),
			formula(2, 2, 15, 7, rgce(ptg.exp(1, 2))),
			formula(3, 2, 15, 11, rgce(ptg.exp(1, 2))),
		];
		const wb = sheetWith(records);
		expect(at(wb, 1, 2)).toEqual({ value: 3, formula: '(B2+A2)*$A$1' });
		expect(at(wb, 2, 2)).toEqual({ value: 7, formula: '(B3+A3)*$A$1' });
		expect(at(wb, 3, 2)).toEqual({ value: 11, formula: '(B4+A4)*$A$1' });
	});

	it('expands shared formulas whose block record follows a string result', () => {
		const wb = sheetWith([
			formula(0, 1, 15, 'a', rgce(ptg.exp(0, 1))),
			shrfmla(0, 1, 1, 1, rgce(ptg.refN(0, -1), ptg.str('!'), ptg.concat)),
			stringResult('a!'),
			formula(1, 1, 15, 'b', rgce(ptg.exp(0, 1))),
			stringResult('b!'),
		]);
		expect(at(wb, 0, 1)).toEqual({ value: 'a!', formula: 'A1&"!"' });
		expect(at(wb, 1, 1)).toEqual({ value: 'b!', formula: 'A2&"!"' });
	});

	it('reads array formulas onto the anchor and links the other cells', () => {
		const tokens = rgce(
			ptg.area(0, 0, 1, 0, { cls: 0x60 }),
			ptg.area(0, 1, 1, 1, { cls: 0x60 }),
			ptg.mul,
		);
		const wb = sheetWith([
			formula(0, 3, 15, 2, rgce(ptg.exp(0, 3))),
			arrayFormula(0, 1, 3, 3, tokens),
			formula(1, 3, 15, 12, rgce(ptg.exp(0, 3))),
		]);
		expect(at(wb, 0, 3)).toEqual({
			value: 2,
			formula: 'A1:A2*B1:B2',
			array: { ref: { top: 0, left: 3, bottom: 1, right: 3 }, dynamic: false },
		});
		expect(at(wb, 1, 3)).toEqual({ value: 12, arrayAnchor: cellKey(0, 3) });
	});

	it('keeps cached values for data tables and missing shared formulas', () => {
		const wb = sheetWith([
			formula(0, 0, 15, 5, rgce(ptg.tbl(0, 0))),
			table(0, 0, 0, 0),
			formula(1, 0, 15, 6, rgce(ptg.exp(9, 9))),
		]);
		expect(at(wb, 0, 0)).toEqual({ value: 5 });
		expect(at(wb, 1, 0)).toEqual({ value: 6 });
		const codes = wb.warnings.map((w) => w.code);
		expect(codes).toContain('data-table');
		expect(codes).toContain('shared-formula-missing');
	});

	it('keeps cached values for unreadable formulas and external references', () => {
		const wb = readXlsWorkbook(
			buildXls(
				[
					supbookSelf(1),
					supbookExternal('\u0001\u0003Other.xls', ['Data']),
					externName('Rate'),
					externSheet([
						[0, 0, 0],
						[1, 0, 0],
					]),
				],
				[
					{
						name: 'Sheet1',
						records: [
							formula(0, 0, 15, 9, rgce(ptg.ref3d(1, 0, 0))),
							formula(1, 0, 15, 4, rgce(ptg.funcVar(250, 1))),
							formula(2, 0, 15, 1, rgce(ptg.ref3d(0, 0, 1))),
						],
					},
				],
			),
		);
		expect(at(wb, 0, 0)).toEqual({ value: 9 });
		expect(at(wb, 1, 0)).toEqual({ value: 4 });
		expect(at(wb, 2, 0)).toEqual({ value: 1, formula: 'Sheet1!B1' });
		const external = wb.warnings.find((w) => w.code === 'external-link');
		expect(external?.message).toContain('Other.xls');
		expect(wb.warnings.find((w) => w.code === 'formula-unreadable')?.cell).toBe('A2');
	});

	it('reads formulas with array constants from the extra data', () => {
		const extra = [0x01, ...u16(0), 0x01, ...new Array<number>(8).fill(0), 0x02, ...xlString('a')];
		const wb = sheetWith(
			[formula(0, 0, 15, 1, rgce(ptg.array(), ptg.funcVar(0, 1)), extra)],
			[font()],
		);
		expect(at(wb, 0, 0)?.formula).toBe('COUNT({0,"a"})');
	});
});

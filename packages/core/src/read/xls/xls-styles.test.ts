import { describe, expect, it } from 'vitest';

import { cellKey } from '../../address/cell-address';
import { readXlsWorkbook } from './index';
import {
	buildXls,
	font,
	format,
	palette,
	styleBuiltin,
	styleUser,
	workbookStream,
	xf,
	xlsFile,
} from './test-utils/biff-builder';
import { blank, number } from './test-utils/sheet-builder';

/** XF indices 0-15 come from the standard styles; extra XFs start at 16 (cellXfs 1). */
function read(globals: number[][], records: number[][] = []) {
	return readXlsWorkbook(buildXls(globals, [{ name: 'Sheet1', records }]));
}

describe('fonts', () => {
	it('converts FONT records, skipping BIFF font index 4', () => {
		const wb = read([
			font({
				name: 'Times New Roman',
				height: 240,
				bold: true,
				italic: true,
				underline: 1,
				color: 10,
			}),
			font({ name: 'Arial', strike: true, underline: 0x22, script: 1, charset: 161, family: 1 }),
			xf({ font: 5 }),
			xf({ font: 6 }),
		]);
		const fonts = wb.styles.fonts;
		expect(fonts).toHaveLength(6);
		expect(fonts[4]).toEqual({
			name: 'Times New Roman',
			size: 12,
			bold: true,
			italic: true,
			underline: 'single',
			color: { indexed: 10 },
			family: 2,
		});
		expect(fonts[5]).toEqual({
			name: 'Arial',
			size: 10,
			strike: true,
			underline: 'doubleAccounting',
			vertAlign: 'superscript',
			family: 1,
			charset: 161,
		});
		expect(wb.styles.cellXfs[1].fontId).toBe(4);
		expect(wb.styles.cellXfs[2].fontId).toBe(5);
	});
});

describe('number formats', () => {
	it('keeps custom formats and localized built-in formats', () => {
		const wb = read([
			format(164, '0.000 "kg"'),
			format(5, '"$"#,##0_);\\("$"#,##0\\)'),
			format(6, '#,##0 "€";[Red]-#,##0 "€"'),
			xf({ numFmt: 164 }),
			xf({ numFmt: 14 }),
		]);
		expect(wb.styles.numFmts).toEqual({ 164: '0.000 "kg"', 6: '#,##0 "€";[Red]-#,##0 "€"' });
		expect(wb.styles.cellXfs[1].numFmtId).toBe(164);
		expect(wb.styles.cellXfs[2].numFmtId).toBe(14);
	});
});

describe('cell formats', () => {
	it('reads alignment, protection and the quote prefix', () => {
		const wb = read([
			xf({ h: 2, v: 0, wrap: true, rotation: 45, indent: 2, shrink: true }),
			xf({ rotation: 255, h: 6, v: 4 }),
			xf({ locked: false, hidden: true, quotePrefix: true }),
		]);
		const [, a, b, c] = wb.styles.cellXfs;
		expect(a.alignment).toEqual({
			horizontal: 'center',
			vertical: 'top',
			wrapText: true,
			textRotation: 45,
			indent: 2,
			shrinkToFit: true,
		});
		expect(b.alignment).toEqual({
			horizontal: 'centerContinuous',
			vertical: 'distributed',
			textRotation: 255,
		});
		expect(c.protection).toEqual({ locked: false, hidden: true });
		expect(c.quotePrefix).toBe(true);
	});

	it('reads borders with styles, colors and diagonals', () => {
		const wb = read([
			xf({
				borders: [
					[1, 10],
					[2, 12],
					[6, 8],
					[13, 64],
				],
				diagonal: { dg: 4, icv: 17, down: true, up: true },
			}),
		]);
		const border = wb.styles.borders[wb.styles.cellXfs[1].borderId];
		expect(border).toEqual({
			left: { style: 'thin', color: { indexed: 10 } },
			right: { style: 'medium', color: { indexed: 12 } },
			top: { style: 'double', color: { indexed: 8 } },
			bottom: { style: 'slantDashDot', color: { indexed: 64 } },
			diagonal: { style: 'dotted', color: { indexed: 17 } },
			diagonalDown: true,
			diagonalUp: true,
		});
	});

	it('reads pattern fills and deduplicates fills and borders', () => {
		const wb = read([
			xf({ fill: { fls: 1, fore: 13, back: 64 } }),
			xf({ fill: { fls: 1, fore: 13, back: 64 }, numFmt: 2 }),
			xf({ fill: { fls: 17, fore: 10, back: 9 } }),
		]);
		const { fills, cellXfs } = wb.styles;
		expect(fills.slice(0, 2)).toEqual([
			{ type: 'pattern', patternType: 'none' },
			{ type: 'pattern', patternType: 'gray125' },
		]);
		expect(fills[cellXfs[1].fillId]).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { indexed: 13 },
			bgColor: { indexed: 64 },
		});
		expect(cellXfs[2].fillId).toBe(cellXfs[1].fillId);
		expect(fills[cellXfs[3].fillId]).toEqual({
			type: 'pattern',
			patternType: 'gray125',
			fgColor: { indexed: 10 },
			bgColor: { indexed: 9 },
		});
		expect(wb.styles.borders).toEqual([{}]);
	});

	it('maps style XFs to cellStyleXfs and cell XFs to cellXfs', () => {
		const wb = read([
			xf({ style: true, font: 1, used: 0xf8 }),
			xf({ parent: 16, font: 1, used: 0x08 }),
			styleBuiltin(16, 3),
			styleBuiltin(1, 1, 1),
			styleUser(16, 'My Style'),
		]);
		const { cellStyleXfs, cellXfs, cellStyles } = wb.styles;
		expect(cellStyleXfs).toHaveLength(16);
		expect(cellXfs).toHaveLength(2);
		expect(cellXfs[1]).toEqual({
			numFmtId: 0,
			fontId: 1,
			fillId: 0,
			borderId: 0,
			applyFont: true,
			xfId: 15,
		});
		expect(cellStyleXfs[15].applyNumberFormat).toBeUndefined();
		expect(cellStyleXfs[15].applyFont).toBe(false);
		expect(cellStyles).toEqual([
			{ name: 'Normal', xfId: 0, builtinId: 0 },
			{ name: 'Comma', xfId: 15, builtinId: 3 },
			{ name: 'RowLevel_2', xfId: 1, builtinId: 1, iLevel: 1 },
			{ name: 'My Style', xfId: 15 },
		]);
	});

	it('gives cells that point at a style XF a derived cell format', () => {
		const wb = read([xf({ style: true, font: 2, h: 3 })], [number(0, 0, 16, 1), blank(0, 1, 16)]);
		const style = wb.sheets[0].cells.get(cellKey(0, 0))?.style;
		expect(style).toBe(1);
		expect(wb.styles.cellXfs[1]).toEqual({
			numFmtId: 0,
			fontId: 2,
			fillId: 0,
			borderId: 0,
			alignment: { horizontal: 'right' },
			xfId: 15,
		});
		expect(wb.sheets[0].cells.get(cellKey(0, 1))?.style).toBe(1);
	});

	it('reads a custom palette into indexedColors', () => {
		const wb = read([palette(['112233', 'AABBCC'])]);
		const colors = wb.styles.indexedColors!;
		expect(colors).toHaveLength(64);
		expect(colors[0]).toBe('FF000000');
		expect(colors[8]).toBe('FF112233');
		expect(colors[9]).toBe('FFAABBCC');
		expect(colors[10]).toBe('FFFF0000');
		expect(read([]).styles.indexedColors).toBeUndefined();
	});

	it('falls back to defaults when the file has no formatting records', () => {
		const wb = readXlsWorkbook(
			xlsFile(workbookStream([], [{ name: 'S', records: [number(0, 0, 0, 1)] }])),
		);
		expect(wb.styles.fonts).toEqual([{ name: 'Arial', size: 10 }]);
		expect(wb.styles.cellXfs).toEqual([
			{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 },
		]);
		expect(wb.styles.cellStyles).toEqual([{ name: 'Normal', xfId: 0, builtinId: 0 }]);
		expect(wb.sheets[0].cells.get(0)).toEqual({ value: 1 });
	});
});

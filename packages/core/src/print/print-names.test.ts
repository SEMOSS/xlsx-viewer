import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import { printColWidth, printRowHeight } from './metrics';
import {
	COMMON_PAPER_SIZES,
	MARGIN_PRESETS,
	marginPresetOf,
	PAPER_SIZES,
	paperDimensions,
	paperSize,
	paperSizeDetail,
} from './paper';
import {
	formatPrintArea,
	formatPrintTitles,
	getPrintArea,
	getPrintTitles,
	parsePrintAreaText,
	parseTitleCols,
	parseTitleRows,
} from './print-names';
import { anchorCellRange, printableRange } from './print-range';

const r = (ref: string) => parseRange(ref)!;

describe('print names', () => {
	it('reads the print area of a sheet (several areas, quoted names)', () => {
		const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet'] });
		wb.definedNames.push(
			{
				name: '_xlnm.Print_Area',
				formula: "'My Sheet'!$A$1:$B$2,'My Sheet'!$D$4",
				localSheet: 1,
				builtin: true,
			},
			{ name: '_xlnm.Print_Area', formula: 'Data!$C$3:$E$9', localSheet: 0, builtin: true },
		);
		expect(getPrintArea(wb, 1)).toEqual([r('A1:B2'), r('D4')]);
		expect(getPrintArea(wb, 0)).toEqual([r('C3:E9')]);
		wb.definedNames.length = 0;
		expect(getPrintArea(wb, 0)).toEqual([]);
	});

	it('reads print titles in either order', () => {
		const wb = createWorkbook();
		wb.definedNames.push({
			name: '_xlnm.Print_Titles',
			formula: 'Sheet1!$1:$2,Sheet1!$B:$C',
			localSheet: 0,
			builtin: true,
		});
		expect(getPrintTitles(wb, 0)).toEqual({
			rows: { start: 0, end: 1 },
			cols: { start: 1, end: 2 },
		});
		wb.definedNames[0].formula = 'Sheet1!$A:$A';
		expect(getPrintTitles(wb, 0)).toEqual({ cols: { start: 0, end: 0 } });
	});

	it('formats print areas and titles like Excel', () => {
		expect(formatPrintArea('Sheet1', [r('A1:D20'), r('F1')])).toBe('Sheet1!$A$1:$D$20,Sheet1!$F$1');
		expect(formatPrintArea('My Sheet', [r('B:C')])).toBe("'My Sheet'!$B:$C");
		expect(
			formatPrintTitles('Sheet1', { rows: { start: 0, end: 1 }, cols: { start: 0, end: 0 } }),
		).toBe('Sheet1!$A:$A,Sheet1!$1:$2');
		expect(formatPrintTitles('Sheet1', {})).toBe('');
	});

	it('parses dialog entries', () => {
		expect(parseTitleRows('$1:$3')).toEqual({ start: 0, end: 2 });
		expect(parseTitleRows('2')).toEqual({ start: 1, end: 1 });
		expect(parseTitleRows("'My Sheet'!$4:$4")).toEqual({ start: 3, end: 3 });
		expect(parseTitleRows('')).toBeUndefined();
		expect(parseTitleRows('A:B')).toBeNull();
		expect(parseTitleRows('nonsense(')).toBeNull();
		expect(parseTitleCols('$A:$B')).toEqual({ start: 0, end: 1 });
		expect(parseTitleCols('c')).toEqual({ start: 2, end: 2 });
		expect(parseTitleCols('1:2')).toBeNull();
		expect(parsePrintAreaText('A1:B2, $D$4', 'Sheet1')).toEqual([r('A1:B2'), r('D4')]);
		expect(parsePrintAreaText('=Sheet1!A1:B2', 'Sheet1')).toEqual([r('A1:B2')]);
		expect(parsePrintAreaText('Other!A1', 'Sheet1')).toBeNull();
		expect(parsePrintAreaText('SUM(A1)', 'Sheet1')).toBeNull();
		expect(parsePrintAreaText(' ', 'Sheet1')).toEqual([]);
	});
});

describe('paper sizes and margins', () => {
	it('knows the common paper sizes in points', () => {
		expect(paperSize(1)).toMatchObject({ name: 'Letter', width: 612, height: 792 });
		expect(paperSize(3)).toMatchObject({ name: 'Tabloid', width: 792, height: 1224 });
		expect(paperSize(5)).toMatchObject({ name: 'Legal', width: 612, height: 1008 });
		expect(paperSize(7)).toMatchObject({ name: 'Executive', width: 522, height: 756 });
		expect(paperSize(8).width).toBeCloseTo(841.89, 1);
		expect(paperSize(9).height).toBeCloseTo(841.89, 1);
		expect(paperSize(11).width).toBeCloseTo(419.53, 1);
		expect(paperSize(13).name).toBe('B5 (JIS)');
		expect(paperSize(999).name).toBe('Letter');
		expect(paperSize(undefined).name).toBe('Letter');
		expect(new Set(PAPER_SIZES.map((p) => p.code)).size).toBe(PAPER_SIZES.length);
		expect(COMMON_PAPER_SIZES.every((c) => paperSize(c).code === c)).toBe(true);
		expect(paperSizeDetail(paperSize(1))).toBe('8.5" x 11"');
		expect(paperSizeDetail(paperSize(9), true)).toBe('21 cm x 29.7 cm');
	});

	it('swaps width and height in landscape', () => {
		expect(paperDimensions({ orientation: 'landscape' })).toEqual({
			width: 792,
			height: 612,
			landscape: true,
		});
		expect(paperDimensions(undefined)).toEqual({ width: 612, height: 792, landscape: false });
	});

	it('recognises margin presets', () => {
		expect(marginPresetOf(undefined)).toBe('normal');
		expect(marginPresetOf(MARGIN_PRESETS.wide)).toBe('wide');
		expect(marginPresetOf({ ...MARGIN_PRESETS.narrow })).toBe('narrow');
		expect(marginPresetOf({ ...MARGIN_PRESETS.narrow, top: 2 })).toBeUndefined();
	});
});

describe('print metrics and the printed range', () => {
	it('measures rows and columns in points', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.rows.set(1, { height: 30 });
		sheet.rows.set(2, { hidden: true, height: 30 });
		sheet.cols = [
			{ min: 1, max: 1, width: 20 },
			{ min: 2, max: 2, width: 20, hidden: true },
		];
		expect(printRowHeight(sheet, 0)).toBe(15);
		expect(printRowHeight(sheet, 1)).toBe(30);
		expect(printRowHeight(sheet, 2)).toBe(0);
		expect(printColWidth(sheet, 0)).toBe(48);
		expect(printColWidth(sheet, 1)).toBe(105);
		expect(printColWidth(sheet, 2)).toBe(0);
		sheet.zeroHeight = true;
		expect(printRowHeight(sheet, 5)).toBe(0);
	});

	it('covers values, visible formatting, merges and objects', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		expect(printableRange(sheet, wb.styles)).toBeNull();
		putCell(sheet, 2, 2, { value: 1 });
		expect(printableRange(sheet, wb.styles)).toEqual(r('C3'));
		wb.styles.fills.push({ type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFF0000' } });
		wb.styles.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: 2, borderId: 0, xfId: 0 });
		wb.styles.cellXfs.push({ numFmtId: 14, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });
		putCell(sheet, 6, 1, { value: null, style: 1 });
		putCell(sheet, 9, 9, { value: null, style: 2 });
		expect(printableRange(sheet, wb.styles)).toEqual(r('B3:C7'));
		sheet.merges = [r('C7:E8')];
		expect(printableRange(sheet, wb.styles)).toEqual(r('B3:E8'));
	});

	it('maps drawing anchors to cells', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		expect(
			anchorCellRange(sheet, {
				type: 'twoCell',
				from: { row: 1, col: 1, rowOff: 0, colOff: 0 },
				to: { row: 4, col: 3, rowOff: 5, colOff: 0 },
			}),
		).toEqual(r('B2:C5'));
		// 100 x 40 pt from B2: two columns (96 pt) and a bit, three rows.
		expect(
			anchorCellRange(sheet, {
				type: 'oneCell',
				from: { row: 1, col: 1, rowOff: 0, colOff: 0 },
				ext: { cx: 100 * 12700, cy: 40 * 12700 },
			}),
		).toEqual(r('B2:D4'));
		expect(
			anchorCellRange(sheet, {
				type: 'absolute',
				pos: { x: 50 * 12700, y: 0 },
				ext: { cx: 10 * 12700, cy: 10 * 12700 },
			}),
		).toEqual(r('B1:B1'));
	});
});

import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { Workbook } from '../types/workbook';
import { setCellInput } from './cell-input';
import {
	type DelimitedSplitOptions,
	guessFixedWidthBreaks,
	parseTextToColumns,
	splitDelimitedText,
	splitFixedWidthText,
	TextToColumnsError,
	textToColumns,
	textToColumnsLines,
	type TextToColumnsOptions,
	textToColumnsOverwrites,
} from './text-to-columns';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;

const NONE: DelimitedSplitOptions = {
	tab: false,
	semicolon: false,
	comma: false,
	space: false,
	treatConsecutiveAsOne: false,
	textQualifier: '"',
};

function column(entries: string[], col = 0): Workbook {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	entries.forEach((v, i) => setCellInput(tx, wb, 0, i, col, v));
	tx.commit();
	return wb;
}

const at = (wb: Workbook, ref: string) => getCell(wb.sheets[0], r(ref).top, r(ref).left);
const row = (wb: Workbook, ref: string) => {
	const a = r(ref);
	const out: unknown[] = [];
	for (let c = a.left; c <= a.right; c++) {
		out.push(getCell(wb.sheets[0], a.top, c)?.value ?? null);
	}
	return out;
};
const fmt = (wb: Workbook, ref: string) =>
	numberFormatCode(wb.styles.cellXfs[at(wb, ref)?.style ?? 0].numFmtId, wb.styles.numFmts);

describe('splitDelimitedText', () => {
	it('splits on the chosen delimiters', () => {
		expect(splitDelimitedText('a,b;c d', { ...NONE, comma: true })).toEqual(['a', 'b;c d']);
		expect(
			splitDelimitedText('a,b;c d', { ...NONE, comma: true, semicolon: true, space: true }),
		).toEqual(['a', 'b', 'c', 'd']);
		expect(splitDelimitedText('a\tb', { ...NONE, tab: true })).toEqual(['a', 'b']);
		expect(splitDelimitedText('a|b|', { ...NONE, other: '|x' })).toEqual(['a', 'b', '']);
		expect(splitDelimitedText('abc', NONE)).toEqual(['abc']);
	});

	it('treats consecutive delimiters as one when asked', () => {
		const opts = { ...NONE, comma: true, space: true };
		expect(splitDelimitedText('a,,b', opts)).toEqual(['a', '', 'b']);
		expect(splitDelimitedText('a, b', { ...opts, treatConsecutiveAsOne: true })).toEqual([
			'a',
			'b',
		]);
		expect(splitDelimitedText(',a', { ...opts, treatConsecutiveAsOne: true })).toEqual(['', 'a']);
	});

	it('keeps delimiters inside qualified text', () => {
		const opts = { ...NONE, comma: true };
		expect(splitDelimitedText('"a,b",c', opts)).toEqual(['a,b', 'c']);
		expect(splitDelimitedText('"say ""hi""",x', opts)).toEqual(['say "hi"', 'x']);
		expect(splitDelimitedText("'a,b',c", { ...opts, textQualifier: "'" })).toEqual(['a,b', 'c']);
		expect(splitDelimitedText('"a,b",c', { ...opts, textQualifier: null })).toEqual([
			'"a',
			'b"',
			'c',
		]);
		expect(splitDelimitedText('x"a,b"', opts)).toEqual(['x"a', 'b"']);
		expect(splitDelimitedText('"open,end', opts)).toEqual(['open,end']);
	});
});

describe('splitFixedWidthText / guessFixedWidthBreaks / parseTextToColumns', () => {
	it('splits at positions and trims trailing spaces', () => {
		expect(splitFixedWidthText('Ann   42  Oslo', [6, 10])).toEqual(['Ann', '42', 'Oslo']);
		expect(splitFixedWidthText('Ann', [6, 10])).toEqual(['Ann']);
		expect(splitFixedWidthText('abcdefgh', [10, 4, 4, 0])).toEqual(['abcd', 'efgh']);
		expect(splitFixedWidthText('  x', [])).toEqual(['  x']);
	});

	it('guesses breaks from columns of spaces shared by every line', () => {
		expect(
			guessFixedWidthBreaks(['Name  Age  City', 'Bob   42   Rome', '', 'Al    7    Oslo']),
		).toEqual([6, 11]);
		expect(guessFixedWidthBreaks(['a b', 'abc'])).toEqual([]);
		expect(guessFixedWidthBreaks(['  lead x'])).toEqual([7]);
		expect(guessFixedWidthBreaks([])).toEqual([]);
	});

	it('parses lines for the preview', () => {
		expect(parseTextToColumns(['a,b', '', 'c'], { mode: 'delimited', comma: true })).toEqual([
			['a', 'b'],
			[],
			['c'],
		]);
		expect(parseTextToColumns(['abcdef'], { mode: 'fixed', breaks: [2] })).toEqual([
			['ab', 'cdef'],
		]);
	});
});

describe('textToColumns', () => {
	const comma = (extra: Partial<TextToColumnsOptions> = {}): TextToColumnsOptions => ({
		mode: 'delimited',
		comma: true,
		destination: { row: 0, col: 0 },
		...extra,
	});

	function convert(wb: Workbook, ref: string, options: TextToColumnsOptions) {
		const tx = new EditTransaction(wb, 'Text to Columns');
		const result = textToColumns(tx, wb, 0, r(ref), options);
		return { result, entry: tx.commit() };
	}

	it('writes General fields like typed input and skips empty rows', () => {
		const wb = column(['Ann,1/2/2024,50%,=1+1', '', 'Bob,"1,5",x']);
		const { result } = convert(wb, 'A1:A3', comma());
		expect(result).toEqual({ columns: 4 });
		expect(row(wb, 'A1:C1')).toEqual(['Ann', 45293, 0.5]);
		expect(fmt(wb, 'B1')).toBe('m/d/yyyy');
		expect(at(wb, 'D1')?.formula).toBe('1+1');
		expect(row(wb, 'A2:D2')).toEqual([null, null, null, null]);
		expect(row(wb, 'A3:D3')).toEqual(['Bob', '1,5', 'x', null]);
	});

	it('stores Text columns as text, skips columns and honours the destination', () => {
		const wb = column(['007,a,b', '12,c,d']);
		convert(
			wb,
			'A1:A2',
			comma({ destination: { row: 0, col: 2 }, columnFormats: ['text', 'skip'] }),
		);
		expect(row(wb, 'C1:E1')).toEqual(['007', 'b', null]);
		expect(fmt(wb, 'C1')).toBe('@');
		expect(row(wb, 'C2:D2')).toEqual(['12', 'd']);
		expect(at(wb, 'A1')?.value).toBe('007,a,b');
	});

	it('converts displayed values of non-text cells', () => {
		const wb = column(['3/8/2006']);
		expect(textToColumnsLines(wb, 0, r('A1:A2'))).toEqual(['3/8/2006', '']);
		convert(wb, 'A1', { mode: 'delimited', other: '/', destination: { row: 0, col: 0 } });
		expect(row(wb, 'A1:C1')).toEqual([3, 8, 2006]);
	});

	it('splits fixed-width text', () => {
		const wb = column(['Ann   42', 'Bo    7']);
		convert(wb, 'A1:A2', { mode: 'fixed', breaks: [6], destination: { row: 0, col: 0 } });
		expect(row(wb, 'A1:B1')).toEqual(['Ann', 42]);
		expect(row(wb, 'A2:B2')).toEqual(['Bo', 7]);
	});

	it('detects data that would be replaced', () => {
		const wb = column(['a,b', 'c']);
		expect(textToColumnsOverwrites(wb, 0, r('A1:A2'), comma())).toBe(false);
		setCellInput(new EditTransaction(wb, 'x'), wb, 0, 1, 1, 'keep');
		// Row 2 has a single field, so B2 is not written.
		expect(textToColumnsOverwrites(wb, 0, r('A1:A2'), comma())).toBe(false);
		setCellInput(new EditTransaction(wb, 'x'), wb, 0, 0, 1, 'old');
		expect(textToColumnsOverwrites(wb, 0, r('A1:A2'), comma())).toBe(true);
		expect(
			textToColumnsOverwrites(wb, 0, r('A1:A2'), comma({ columnFormats: ['general', 'skip'] })),
		).toBe(false);
	});

	it('refuses several columns, merged cells and legacy arrays without changing anything', () => {
		const wb = column(['a,b']);
		const reason = (fn: () => void) => {
			try {
				fn();
			} catch (e) {
				return e instanceof TextToColumnsError ? e.reason : 'other';
			}
			return null;
		};
		expect(reason(() => convert(wb, 'A1:B1', comma()))).toBe('multipleColumns');
		expect(new TextToColumnsError('x', 'merge').name).toBe('TextToColumnsError');
		wb.sheets[0].merges.push(r('B1:C1'));
		expect(reason(() => convert(wb, 'A1', comma()))).toBe('merge');
		wb.sheets[0].merges = [];
		wb.sheets[0].cells.set(cellKey(0, 2), {
			value: 1,
			formula: 'X',
			array: { ref: r('C1:D1'), dynamic: false },
		});
		wb.sheets[0].cells.set(cellKey(0, 1), { value: 1, arrayAnchor: cellKey(0, 2) });
		expect(reason(() => convert(wb, 'A1', comma()))).toBe('array');
		expect(at(wb, 'A1')?.value).toBe('a,b');
	});

	it('undoes and redoes', () => {
		const wb = column(['x,1', 'y,2']);
		const { entry } = convert(wb, 'A1:A2', comma());
		expect(row(wb, 'A1:B2')).toEqual(['x', 1]);
		entry.undo();
		expect(row(wb, 'A1:B1')).toEqual(['x,1', null]);
		expect(row(wb, 'A2:B2')).toEqual(['y,2', null]);
		entry.redo();
		expect(row(wb, 'A2:B2')).toEqual(['y', 2]);
	});
});

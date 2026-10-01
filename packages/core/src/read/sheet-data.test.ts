import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import type { Cell } from '../types/cell';
import { ERRORS } from '../values/errors';
import { isoToSerial } from './cell-values';
import type { SharedStrings } from './context';
import { parseSheetData, type SheetDataOptions, type SheetDataResult } from './sheet-data';

const SST: SharedStrings = {
	strings: ['hello', 'rich', 'world'],
	rich: new Map([[1, [{ text: 'ri', font: { bold: true } }, { text: 'ch' }]]]),
};

function scan(content: string, opts: Partial<SheetDataOptions> = {}): SheetDataResult {
	const p = opts.prefix ? `${opts.prefix}:` : '';
	const open = `<${p}sheetData>`;
	const xml = `${open}${content}</${p}sheetData>`;
	return parseSheetData(xml, open.length, xml.length - open.length - 1, {
		prefix: '',
		sharedStrings: SST,
		date1904: false,
		...opts,
	});
}

function at(result: SheetDataResult, ref: string): Cell | undefined {
	const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
	let col = 0;
	for (const ch of m[1]) {
		col = col * 26 + (ch.charCodeAt(0) - 64);
	}
	return result.cells.get(cellKey(Number(m[2]) - 1, col - 1));
}

describe('parseSheetData: values', () => {
	it('reads every cell type', () => {
		const r = scan(
			'<row r="1">' +
				'<c r="A1"><v>42.5</v></c>' +
				'<c r="B1" t="n"><v>7.0000000000000007E-2</v></c>' +
				'<c r="C1" t="s"><v>0</v></c>' +
				'<c r="D1" t="s"><v>1</v></c>' +
				'<c r="E1" t="str"><v>a &amp; b_x000D_</v></c>' +
				'<c r="F1" t="b"><v>1</v></c>' +
				'<c r="G1" t="b"><v>0</v></c>' +
				'<c r="H1" t="e"><v>#DIV/0!</v></c>' +
				'<c r="I1" t="e"><v>#PYTHON!</v></c>' +
				'<c r="J1" t="inlineStr"><is><t>inline &lt;text&gt;</t></is></c>' +
				'<c r="K1" t="d"><v>2024-01-15T12:00:00Z</v></c>' +
				'<c r="L1" s="3"/>' +
				'<c r="M1"><v></v></c>' +
				'<c r="N1"><v/></c>' +
				'<c r="O1"><v>abc</v></c>' +
				'<c r="P1" t="b"><v>true</v></c>' +
				'</row>',
		);
		expect(at(r, 'A1')).toEqual({ value: 42.5 });
		expect(at(r, 'B1')?.value).toBeCloseTo(0.07, 15);
		expect(at(r, 'C1')).toEqual({ value: 'hello' });
		expect(at(r, 'D1')).toEqual({ value: 'rich', richText: SST.rich.get(1) });
		expect(at(r, 'E1')?.value).toBe('a & b\r');
		expect(at(r, 'F1')?.value).toBe(true);
		expect(at(r, 'G1')?.value).toBe(false);
		expect(at(r, 'H1')?.value).toBe(ERRORS.DIV0);
		expect(at(r, 'I1')?.value).toBe(ERRORS.UNKNOWN);
		expect(at(r, 'J1')).toEqual({ value: 'inline <text>' });
		expect(at(r, 'K1')?.value).toBe(45306.5);
		expect(at(r, 'L1')).toEqual({ value: null, style: 3 });
		expect(at(r, 'M1')).toEqual({ value: null });
		expect(at(r, 'N1')).toEqual({ value: null });
		expect(at(r, 'O1')?.value).toBe('abc');
		expect(at(r, 'P1')?.value).toBe(true);
		expect(r.issues).toEqual([]);
	});

	it('reads rich inline strings and ignores phonetic runs', () => {
		const r = scan(
			'<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/><sz val="12"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>Bold</t></r>' +
				'<r><t xml:space="preserve"> plain</t></r><rPh sb="0" eb="1"><t>PH</t></rPh></is></c></row>',
		);
		expect(at(r, 'A1')).toEqual({
			value: 'Bold plain',
			richText: [
				{ text: 'Bold', font: { bold: true, size: 12, color: { rgb: 'FFFF0000' }, name: 'Arial' } },
				{ text: ' plain' },
			],
		});
	});

	it('keeps cm, vm and ph and omits style 0', () => {
		const r = scan('<row r="2"><c r="B2" s="0" cm="1" vm="2" ph="1"><v>1</v></c></row>');
		expect(at(r, 'B2')).toEqual({ value: 1, cm: 1, vm: 2, ph: true });
	});

	it('flags missing shared strings', () => {
		const r = scan('<row r="1"><c r="A1" t="s"><v>99</v></c><c r="A2" t="s"><v>98</v></c></row>');
		expect(at(r, 'A1')?.value).toBe('');
		expect(r.issues).toHaveLength(1);
		expect(r.issues[0].code).toBe('invalid-shared-string');
	});

	it('reads ISO dates in the 1904 system', () => {
		const r = scan(
			'<row r="1"><c r="A1" t="d"><v>1904-01-02</v></c><c r="B1" t="d"><v>not a date</v></c></row>',
			{
				date1904: true,
			},
		);
		expect(at(r, 'A1')?.value).toBe(1);
		expect(at(r, 'B1')?.value).toBe('not a date');
	});
});

describe('parseSheetData: positions and rows', () => {
	it('infers missing row and cell references', () => {
		const r = scan(
			'<row><c><v>1</v></c><c><v>2</v></c><c r="E1"><v>5</v></c><c><v>6</v></c></row>' +
				'<row><c><v>10</v></c></row>' +
				'<row r="5"><c><v>50</v></c></row>' +
				'<row><c t="s"><v>2</v></c></row>',
		);
		expect(at(r, 'A1')?.value).toBe(1);
		expect(at(r, 'B1')?.value).toBe(2);
		expect(at(r, 'E1')?.value).toBe(5);
		expect(at(r, 'F1')?.value).toBe(6);
		expect(at(r, 'A2')?.value).toBe(10);
		expect(at(r, 'A5')?.value).toBe(50);
		expect(at(r, 'A6')?.value).toBe('world');
	});

	it('reads row attributes and only keeps meaningful row records', () => {
		const r = scan(
			'<row r="1" spans="1:3" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row>' +
				'<row r="2" ht="30" customHeight="1" hidden="1" s="4" customFormat="1" outlineLevel="2" collapsed="1" thickTop="1" thickBot="true"/>' +
				'<row r="3" ht="12.75"></row>',
		);
		expect(r.rows.has(0)).toBe(false);
		expect(r.rows.get(1)).toEqual({
			height: 30,
			customHeight: true,
			hidden: true,
			style: 4,
			customFormat: true,
			outlineLevel: 2,
			collapsed: true,
			thickTop: true,
			thickBot: true,
		});
		expect(r.rows.get(2)).toEqual({ height: 12.75 });
	});

	it('handles pretty-printed XML, comments, CDATA and extension elements', () => {
		const r = scan(
			'\n  <!-- a comment with <c r="Z9"> inside -->\n  <row r="1">\n    <c r="A1" t="str">\n      <f>"x"</f>\n      <v>x</v>\n    </c>\n' +
				'    <extLst><ext uri="{X}"><c r="Z1"><v>9</v></c></ext></extLst>\n  </row>\n',
		);
		expect(r.cells.size).toBe(1);
		expect(at(r, 'A1')).toEqual({ value: 'x', formula: '"x"' });
	});

	it('reads sheetData whose SpreadsheetML elements use a prefix', () => {
		const r = scan(
			'<x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:f>A1&amp;"!"</x:f><x:v>3</x:v></x:c>' +
				'<x:c r="C1" t="inlineStr"><x:is><x:r><x:rPr><x:i/></x:rPr><x:t>it</x:t></x:r></x:is></x:c></x:row>',
			{ prefix: 'x' },
		);
		expect(at(r, 'A1')?.value).toBe('hello');
		expect(at(r, 'B1')).toEqual({ value: 3, formula: 'A1&"!"' });
		expect(at(r, 'C1')).toEqual({
			value: 'it',
			richText: [{ text: 'it', font: { italic: true } }],
		});
	});
});

describe('parseSheetData: formulas', () => {
	it('strips future-function prefixes and decodes entities', () => {
		const r = scan(
			'<row r="1"><c r="A1"><f>_xlfn.XLOOKUP(B1,C:C,D:D)&amp;"&lt;&gt;"</f><v>0</v></c>' +
				'<c r="B1"><f>_xlfn._xlws.FILTER(C1:C9,C1:C9&gt;0)</f></c></row>',
		);
		expect(at(r, 'A1')?.formula).toBe('XLOOKUP(B1,C:C,D:D)&"<>"');
		expect(at(r, 'B1')).toEqual({ value: null, formula: 'FILTER(C1:C9,C1:C9>0)' });
	});

	it('expands shared formulas relative to the master cell', () => {
		const r = scan(
			'<row r="1"><c r="B1"><f t="shared" ref="B1:C3" si="0">A1*2+$A$1</f><v>2</v></c><c r="C1"><f t="shared" si="0"/><v>4</v></c></row>' +
				'<row r="2"><c r="B2"><f t="shared" si="0"/><v>6</v></c><c r="C2"><f t="shared" si="0"/></c></row>' +
				'<row r="3"><c r="B3"><f t="shared" si="0"/></c><c r="C3"><f t="shared" si="0"/><v>8</v></c></row>',
		);
		expect(at(r, 'B1')).toEqual({ value: 2, formula: 'A1*2+$A$1' });
		expect(at(r, 'C1')).toEqual({ value: 4, formula: 'B1*2+$A$1' });
		expect(at(r, 'B2')?.formula).toBe('A2*2+$A$1');
		expect(at(r, 'C2')?.formula).toBe('B2*2+$A$1');
		expect(at(r, 'B3')?.formula).toBe('A3*2+$A$1');
		expect(at(r, 'C3')).toEqual({ value: 8, formula: 'B3*2+$A$1' });
	});

	it('resolves shared-formula members that precede their master and reports orphans', () => {
		const r = scan(
			'<row r="1"><c r="A1"><f t="shared" si="3"/><v>1</v></c></row>' +
				'<row r="2"><c r="A2"><f t="shared" ref="A1:A2" si="3">B2+1</f><v>2</v></c></row>' +
				'<row r="3"><c r="A3"><f t="shared" si="7"/><v>5</v></c></row>',
		);
		expect(at(r, 'A1')?.formula).toBe('B1+1');
		expect(at(r, 'A3')).toEqual({ value: 5 });
		expect(r.issues.map((i) => i.code)).toEqual(['shared-formula-missing']);
	});

	it('records legacy array formulas and links the cells of the result', () => {
		const r = scan(
			'<row r="1"><c r="C1"><f t="array" ref="C1:C3">A1:A3*B1:B3</f><v>1</v></c></row>' +
				'<row r="2"><c r="C2"><v>4</v></c><c r="D2"><v>0</v></c></row>' +
				'<row r="3"><c r="C3"><v>9</v></c></row>',
		);
		expect(at(r, 'C1')).toEqual({
			value: 1,
			formula: 'A1:A3*B1:B3',
			array: { ref: { top: 0, left: 2, bottom: 2, right: 2 }, dynamic: false },
		});
		const anchor = cellKey(0, 2);
		expect(at(r, 'C2')).toEqual({ value: 4, arrayAnchor: anchor });
		expect(at(r, 'C3')).toEqual({ value: 9, arrayAnchor: anchor });
		expect(at(r, 'D2')).toEqual({ value: 0 });
		expect(r.cells.has(cellKey(3, 2))).toBe(false);
	});

	it('marks dynamic arrays through cell metadata', () => {
		const r = scan(
			'<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">_xlfn.SEQUENCE(2)</f><v>1</v></c>' +
				'<c r="B1" cm="2"><f t="array" ref="B1">_xlfn.SINGLE(A1)</f><v>1</v></c></row>' +
				'<row r="2"><c r="A2"><v>2</v></c></row>',
			{ dynamicArrayCm: new Set([1]) },
		);
		expect(at(r, 'A1')).toEqual({
			value: 1,
			formula: 'SEQUENCE(2)',
			array: { ref: { top: 0, left: 0, bottom: 1, right: 0 }, dynamic: true },
			cm: 1,
		});
		expect(at(r, 'A2')?.arrayAnchor).toBe(cellKey(0, 0));
		expect(at(r, 'B1')?.array).toEqual({
			ref: { top: 0, left: 1, bottom: 0, right: 1 },
			dynamic: false,
		});
	});

	it('keeps cached values of what-if data tables and counts them', () => {
		const r = scan(
			'<row r="2"><c r="B2"><f t="dataTable" ref="B2:B3" dt2D="0" dtr="0" r1="A1"/><v>10</v></c></row>' +
				'<row r="3"><c r="B3"><v>20</v></c></row>',
		);
		expect(at(r, 'B2')).toEqual({ value: 10 });
		expect(r.dataTableCells).toBe(1);
	});
});

describe('isoToSerial', () => {
	it.each([
		['2024-01-15', 45306],
		['2024-01-15T06:00:00', 45306.25],
		['2024-01-15T06:00:00.500+02:00', 45306.25 + 0.5 / 86400],
		['1900-01-01', 1],
		['1900-03-01', 61],
		['T12:00:00', 0.5],
		['18:00', 0.75],
	])('%s -> %d', (text, serial) => {
		expect(isoToSerial(text, false)).toBeCloseTo(serial, 9);
	});

	it('rejects non-dates', () => {
		expect(isoToSerial('', false)).toBeUndefined();
		expect(isoToSerial('2024-13-01', false)).toBeUndefined();
		expect(isoToSerial('hello', false)).toBeUndefined();
	});
});

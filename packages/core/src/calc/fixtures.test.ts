import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { readWorkbook } from '../read';
import { buildWorkbook, worksheet } from '../read/test-utils/package-builder';
import { compareWithCachedValues, type Comparison } from './test-utils/compare';

const FIXTURES = join(__dirname, '../__tests__/fixtures/real');

function report(name: string, c: Comparison): string {
	return [
		`${name}: ${c.compared} compared, ${c.skipped} skipped`,
		...c.mismatches.map(
			(m) =>
				`  ${m.sheet}!${m.cell} =${m.formula}: Excel ${JSON.stringify(m.cached)}, engine ${JSON.stringify(m.computed)}`,
		),
	].join('\n');
}

describe('real fixtures reproduce Excel cached values', () => {
	const files = readdirSync(FIXTURES).filter((f) => f.endsWith('.xlsx'));

	it('has fixtures', () => {
		expect(files.length).toBeGreaterThan(0);
	});

	it.each(files)('%s', async (file) => {
		const wb = await readWorkbook(readFileSync(join(FIXTURES, file)));
		const c = compareWithCachedValues(wb);
		expect(c.mismatches, report(file, c)).toEqual([]);
	});
});

const cell = (ref: string, inner: string, attrs = '') => `<c r="${ref}"${attrs}>${inner}</c>`;
const num = (ref: string, v: number) => cell(ref, `<v>${v}</v>`);
const str = (ref: string, sst: number) => cell(ref, `<v>${sst}</v>`, ' t="s"');
const f = (ref: string, formula: string, v: string, attrs = '', fAttrs = '') =>
	cell(
		ref,
		`<f${fAttrs}>${formula.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</f><v>${v}</v>`,
		attrs,
	);

/** A workbook written the way Excel writes it, with Excel's cached results. */
async function goldenWorkbook(): Promise<Uint8Array> {
	const words = ['apple', 'banana', 'cherry', 'date', 'elder'];
	const dataRows = [1, 2, 3, 4, 5].map((n, i) => {
		const r = n;
		const shared = i === 0 ? ' t="shared" ref="C1:C5" si="0"' : ' t="shared" si="0"';
		const c =
			i === 0
				? f(`C${r}`, 'A1*2', String(n * 2), '', shared)
				: cell(`C${r}`, `<f${shared}/><v>${n * 2}</v>`);
		const extra =
			i === 0
				? [
						f('D1', 'SUM(A1:A5)', '15'),
						f('E1', 'VLOOKUP(3,A1:B5,2,FALSE)', 'cherry', ' t="str"'),
						f('F1', 'IF(AND(A1>0,A2>0),"both","no")', 'both', ' t="str"'),
						f('G1', 'AVERAGE(A1:A5)', '3'),
						f('H1', 'ROUND(PI(),2)', '3.14'),
						f('I1', 'LEN(B1)&"-"&UPPER(B2)', '5-BANANA', ' t="str"'),
						f('J1', 'COUNTIF(A1:A5,">2")', '3'),
						f('K1', 'SUMIF(A1:A5,">2")', '12'),
						f('L1', 'INDEX(B1:B5,MATCH(4,A1:A5,0))', 'date', ' t="str"'),
						f('M1', 'Total*2', '30'),
						f('N1', 'DATE(2024,1,15)', '45306'),
						f('O1', 'MOD(17,5)+INT(7.8)', '9'),
						f('P1', '1/0', '#DIV/0!', ' t="e"'),
						f('Q1', 'ISERROR(P1)', '1', ' t="b"'),
						f('R1', 'MAX(A1:A5)-MIN(A1:A5)', '4'),
						f('S1', 'SUMPRODUCT(A1:A5,A1:A5)', '55'),
					].join('')
				: '';
		return `<row r="${r}">${num(`A${r}`, n)}${str(`B${r}`, i)}${c}${extra}</row>`;
	});
	const calcRows = [
		`<row r="1">${f('A1', 'SUM(Data!A1:A5*Data!A1:A5)', '55', '', ' t="array" ref="A1"')}` +
			`${f('B1', '_xlfn.SEQUENCE(3)', '1', ' cm="1"', ' t="array" ref="B1:B3"')}` +
			`${f('C1', 'SUM(_xlfn.ANCHORARRAY(B1))', '6')}${f('D1', 'Data!D1+Data!M1', '45')}` +
			`${f('E1', 'Data!A1:A3*2', '2', '', ' t="array" ref="E1:E3"')}` +
			`${f('F1', '_xlfn.XLOOKUP("date",Data!B1:B5,Data!A1:A5)', '4')}` +
			`${f('G1', '_xlfn.LET(_xlpm.x,10,_xlpm.x*Data!A2)', '20')}` +
			`${f('H1', '_xlfn.UNKNOWNFN(1)', '7')}${f('I1', 'CONCATENATE("a","b")', 'ab', ' t="str"')}</row>`,
		`<row r="2">${num('B2', 2)}${num('E2', 4)}</row>`,
		`<row r="3">${num('B3', 3)}${num('E3', 6)}</row>`,
	];
	const metadata =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
		'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
		'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" ' +
		'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" ' +
		'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
		'<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
		'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
		'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
	return buildWorkbook({
		sheets: [
			{ name: 'Data', xml: worksheet(dataRows.join('')) },
			{ name: 'Calc', xml: worksheet(calcRows.join('')) },
		],
		sharedStrings: words,
		workbookAfterSheets:
			'<definedNames><definedName name="Total">SUM(Data!$A$1:$A$5)</definedName></definedNames>',
		workbookRels: [{ id: 'rIdMeta', type: 'sheetMetadata', target: 'metadata.xml' }],
		parts: { 'xl/metadata.xml': metadata },
	});
}

describe('a workbook as saved by Excel', () => {
	it('recalculates to the cached values', async () => {
		const wb = await readWorkbook(await goldenWorkbook());
		const c = compareWithCachedValues(wb);
		expect(c.mismatches, report('golden', c)).toEqual([]);
		// Everything but UNKNOWNFN (and its absence from the registry) is compared.
		expect(c.compared).toBeGreaterThanOrEqual(25);
		expect(c.skipped).toBeGreaterThanOrEqual(1);
	});

	it('keeps the unsupported formula stale and the dynamic array spilled', async () => {
		const wb = await readWorkbook(await goldenWorkbook());
		compareWithCachedValues(wb);
		const calc = wb.sheets[1];
		expect(calc.cells.get(7)).toMatchObject({ value: 7, stale: true });
		expect(calc.cells.get(1)?.array).toEqual({
			ref: { top: 0, left: 1, bottom: 2, right: 1 },
			dynamic: true,
		});
		expect(
			wb.warnings.some((w) => w.code === 'unsupported-function' && w.message.includes('UNKNOWNFN')),
		).toBe(true);
	});
});

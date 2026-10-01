import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { columnToLetters } from '../address/cell-address';
import { readWorkbook } from './read-workbook';
import { parseSharedStrings } from './shared-strings';
import { parseSheetData } from './sheet-data';
import { MAIN_NS, buildWorkbook, worksheet } from './test-utils/package-builder';

/** A realistic sheetData body: numbers, shared strings, styled cells and a shared formula column. */
function generateSheetData(rows: number, cols: number): string {
	const out: string[] = [];
	for (let r = 1; r <= rows; r++) {
		let row = `<row r="${r}" spans="1:${cols}" x14ac:dyDescent="0.25">`;
		for (let c = 0; c < cols; c++) {
			const ref = `${columnToLetters(c)}${r}`;
			if (c === 0) {
				row += `<c r="${ref}" t="s"><v>${r % 100}</v></c>`;
			} else if (c === cols - 1) {
				row +=
					r === 1
						? `<c r="${ref}"><f t="shared" ref="${ref}:${columnToLetters(c)}${rows}" si="0">B1*2</f><v>${r * 2}</v></c>`
						: `<c r="${ref}"><f t="shared" si="0"/><v>${r * 2}</v></c>`;
			} else if (c % 3 === 0) {
				row += `<c r="${ref}" s="2"><v>${(r * c) / 7}</v></c>`;
			} else {
				row += `<c r="${ref}"><v>${r * c}</v></c>`;
			}
		}
		out.push(`${row}</row>`);
	}
	return out.join('');
}

const SST = { strings: Array.from({ length: 100 }, (_, i) => `label ${i}`), rich: new Map() };

describe('shared strings performance', () => {
	it('reads 200k shared strings (10% rich) quickly', () => {
		const items: string[] = [];
		for (let i = 0; i < 200_000; i++) {
			items.push(
				i % 10 === 0
					? `<si><r><rPr><b/><sz val="11"/></rPr><t>Bold ${i}</t></r><r><t xml:space="preserve"> tail</t></r></si>`
					: `<si><t>Item number ${i} &amp; more</t></si>`,
			);
		}
		const xml = `<sst xmlns="${MAIN_NS}" count="200000" uniqueCount="200000">${items.join('')}</sst>`;
		const t0 = performance.now();
		const table = parseSharedStrings(xml);
		const ms = performance.now() - t0;
		expect(table.strings).toHaveLength(200_000);
		expect(table.strings[1]).toBe('Item number 1 & more');
		expect(table.rich.size).toBe(20_000);
		console.info(`[perf] parseSharedStrings: 200000 items in ${ms.toFixed(1)} ms`);
		expect(ms).toBeLessThan(2000);
	});
});

describe('sheetData performance', () => {
	it('scans 200k cells well within budget', () => {
		const rows = 20_000;
		const cols = 10;
		const xml = `<sheetData>${generateSheetData(rows, cols)}</sheetData>`;
		const opts = { prefix: '', sharedStrings: SST, date1904: false };
		parseSheetData(xml, 11, xml.length - 12, opts); // warm-up
		const t0 = performance.now();
		const result = parseSheetData(xml, 11, xml.length - 12, opts);
		const ms = performance.now() - t0;
		expect(result.cells.size).toBe(rows * cols);
		expect(result.cells.get(cellKey(9_999, 9))?.formula).toBe('B10000*2');
		expect(result.cells.get(cellKey(0, 0))?.value).toBe('label 1');
		console.info(`[perf] parseSheetData: ${rows * cols} cells in ${ms.toFixed(1)} ms`);
		// Target is < 600 ms for 200k cells (100k < 300 ms); generous margin for slow CI machines.
		expect(ms).toBeLessThan(1500);
	});

	it('loads a 200k-cell workbook end to end', async () => {
		const rows = 20_000;
		const bytes = await buildWorkbook({
			sheets: [{ name: 'Big', xml: worksheet(generateSheetData(rows, 10)) }],
			sharedStrings: SST.strings,
		});
		const t0 = performance.now();
		const wb = await readWorkbook(bytes);
		const ms = performance.now() - t0;
		expect(wb.sheets[0].cells.size).toBe(rows * 10);
		console.info(
			`[perf] readWorkbook: 200000 cells in ${ms.toFixed(1)} ms (${bytes.length} bytes zipped)`,
		);
		expect(ms).toBeLessThan(4000);
	}, 30_000);
});

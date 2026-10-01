import { describe, expect, it } from 'vitest';

import { cellKey } from '../../address/cell-address';
import { readXlsWorkbook } from './index';
import { sst, standardStyles, workbookStream, xlsFile } from './test-utils/biff-builder';
import { ptg, rgce } from './test-utils/ptg-builder';
import { formula, labelSst, number, shrfmla } from './test-utils/sheet-builder';

/** Rows of a number, a shared string, a shared formula and a plain formula. */
function bigSheet(rows: number): number[][] {
	const shared = rgce(
		ptg.refN(0, -2),
		ptg.refN(0, -1),
		ptg.mul,
		ptg.area(0, 0, 9, 0),
		ptg.attrSum,
		ptg.add,
	);
	const plain = (r: number): Uint8Array =>
		rgce(ptg.ref(r, 0), ptg.ref(r, 2), ptg.add, ptg.int(2), ptg.func(27));
	const records: number[][] = [];
	for (let r = 0; r < rows; r++) {
		const block = Math.floor(r / 1000) * 1000;
		records.push(number(r, 0, 15, r), labelSst(r, 1, 15, r % 3));
		records.push(formula(r, 2, 15, r * 2, rgce(ptg.exp(block, 2))));
		if (r === block) {
			records.push(shrfmla(block, Math.min(block + 999, rows - 1), 2, 2, shared));
		}
		records.push(formula(r, 3, 15, r, plain(r)));
	}
	return records;
}

describe('.xls performance', () => {
	it('reads 200k cells with 100k formulas well within budget', () => {
		const rows = 50_000;
		const stream = workbookStream(
			[...standardStyles(), sst(['a', 'b', 'c'])],
			[{ name: 'Big', records: bigSheet(rows) }],
		);
		const bytes = xlsFile(stream);
		const t0 = performance.now();
		const wb = readXlsWorkbook(bytes);
		const ms = performance.now() - t0;
		const cells = wb.sheets[0].cells;
		expect(cells.size).toBe(rows * 4);
		expect(cells.get(cellKey(49_999, 2))?.formula).toBe('A50000*B50000+SUM(A1:A10)');
		expect(cells.get(cellKey(5, 3))?.formula).toBe('ROUND(A6+C6,2)');
		expect(wb.warnings).toEqual([]);
		console.info(`[perf] readXlsWorkbook: ${rows * 4} cells in ${ms.toFixed(1)} ms`);
		// About 0.5 s on a laptop; generous margin for slow CI machines.
		expect(ms).toBeLessThan(5000);
	}, 30_000);
});

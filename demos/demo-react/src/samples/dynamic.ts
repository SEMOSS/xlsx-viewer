import type { RangeAddress, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	parseRange,
	setCellInput,
	setColumnWidth,
} from 'xlsx-viewer-core';

/** Dynamic arrays and modern functions: FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, LET and LAMBDA. */
export function buildDynamicArraysWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Dynamic arrays'], title: 'Dynamic Arrays' });
	const tx = new EditTransaction(wb, 'sample');
	const r = (a1: string): RangeAddress => parseRange(a1)!;
	const put = (a1: string, text: string) => {
		const ref = r(a1);
		setCellInput(tx, wb, 0, ref.top, ref.left, text);
	};
	const style = (a1: string, patch: StylePatch) => applyStylePatch(tx, wb, 0, [r(a1)], patch);
	const heading = (a1: string, text: string) => {
		put(a1, text);
		style(a1, { font: { bold: true, color: { rgb: 'FF107C41' } } });
	};

	const data: Array<[string, string, number]> = [
		['Ana', 'North', 420],
		['Ben', 'South', 310],
		['Chen', 'North', 515],
		['Dana', 'East', 280],
		['Eli', 'South', 455],
		['Fay', 'East', 390],
		['Gus', 'North', 205],
	];
	['Name', 'Region', 'Sales'].forEach((h, i) => setCellInput(tx, wb, 0, 0, i, h));
	style('A1:C1', {
		font: { bold: true },
		fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFD9EAD3' } },
	});
	data.forEach(([n, reg, v], i) => {
		setCellInput(tx, wb, 0, i + 1, 0, n);
		setCellInput(tx, wb, 0, i + 1, 1, reg);
		setCellInput(tx, wb, 0, i + 1, 2, String(v));
	});

	heading('E1', 'SORT by sales, descending');
	put('E2', '=SORT(A2:C8,3,-1)');
	heading('I1', 'FILTER: North only');
	put('I2', '=FILTER(A2:C8,B2:B8="North")');
	heading('E11', 'UNIQUE regions and their totals');
	put('E12', '=UNIQUE(B2:B8)');
	put('F12', '=SUMIF(B2:B8,E12#,C2:C8)');
	heading('I11', 'SEQUENCE: a 4 x 3 grid');
	put('I12', '=SEQUENCE(4,3,10,10)');
	heading('A11', 'XLOOKUP');
	put('A12', 'Chen');
	put('B12', '=XLOOKUP(A12,A2:A8,C2:C8,"not found")');
	heading('A15', 'LET and LAMBDA');
	put('A16', 'Average above 300');
	put('B16', '=LET(s,C2:C8,AVERAGE(FILTER(s,s>300)))');
	put('A17', 'Bonus (LAMBDA)');
	put('B17', '=LAMBDA(x,IF(x>400,x*0.1,0))(C4)');
	put('A18', 'Top seller');
	put('B18', '=INDEX(SORTBY(A2:A8,C2:C8,-1),1)');
	style('B16:B17', { numFmt: '0.0' });
	put('A20', 'Edit a sales figure: every spilled range updates.');
	style('A20', { font: { italic: true, color: { rgb: 'FF7F7F7F' } } });

	setColumnWidth(tx, wb, 0, 0, 0, 21);
	setColumnWidth(tx, wb, 0, 4, 4, 12);
	setColumnWidth(tx, wb, 0, 8, 8, 10);
	tx.commit();
	return wb;
}

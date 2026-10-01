import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { paginateSheet } from './paginate';
import type { PrintLayout } from './types';

/**
 * Letter portrait with Normal margins: 511.2 x 684 pt printable. Default
 * columns are 64 px = 48 pt (10 per page), default rows 15 pt (45 per page).
 */

function book(lastCell = 'T100', name = 'Sheet1'): Workbook {
	const wb = createWorkbook({ sheetNames: [name] });
	const sheet = wb.sheets[0];
	const r = parseRange(lastCell)!;
	putCell(sheet, 0, 0, { value: 'first' });
	putCell(sheet, r.bottom, r.right, { value: 'last' });
	return wb;
}

const bodies = (layout: PrintLayout) =>
	layout.pages.map((p) => {
		const b = p.body;
		return `${b.top}-${b.bottom}/${b.left}-${b.right}`;
	});

const range = (ref: string): RangeAddress => parseRange(ref)!;

describe('paginateSheet: paper, margins and automatic breaks', () => {
	it('breaks the used range down, then over', () => {
		const layout = paginateSheet(book(), 0);
		expect(layout.paper).toEqual({ width: 612, height: 792, paperSize: 1, landscape: false });
		expect(layout.margins).toEqual({
			left: 50.4,
			right: 50.4,
			top: 54,
			bottom: 54,
			header: 21.6,
			footer: 21.6,
		});
		expect(layout.printable.width).toBeCloseTo(511.2);
		expect(layout.printable.height).toBe(684);
		expect(layout.scale).toBe(1);
		expect(bodies(layout)).toEqual([
			'0-44/0-9',
			'45-89/0-9',
			'90-99/0-9',
			'0-44/10-19',
			'45-89/10-19',
			'90-99/10-19',
		]);
		expect(layout.pages.map((p) => p.pageNumber)).toEqual([1, 2, 3, 4, 5, 6]);
		expect(layout.pageCount).toBe(6);
		expect(layout.totalPages).toBe(6);
		expect(layout.rowBreaks).toEqual([45, 90]);
		expect(layout.colBreaks).toEqual([10]);
	});

	it('orders pages over, then down', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { pageOrder: 'overThenDown' };
		expect(bodies(paginateSheet(wb, 0)).slice(0, 3)).toEqual([
			'0-44/0-9',
			'0-44/10-19',
			'45-89/0-9',
		]);
	});

	it('uses the paper size and orientation', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { paperSize: 9, orientation: 'landscape' };
		const layout = paginateSheet(wb, 0);
		expect(layout.paper.width).toBeCloseTo(841.89, 1);
		expect(layout.paper.height).toBeCloseTo(595.28, 1);
		expect(layout.paper.landscape).toBe(true);
		// 741.09 pt wide: 15 columns; 487.28 pt tall: 32 rows.
		expect(bodies(layout)[0]).toBe('0-31/0-14');
	});

	it('uses custom margins', () => {
		const wb = book();
		wb.sheets[0].pageMargins = { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 };
		const layout = paginateSheet(wb, 0);
		expect(layout.printable).toEqual({ x: 72, y: 72, width: 468, height: 648 });
		expect(bodies(layout)[0]).toBe('0-42/0-8');
	});

	it('skips hidden rows and columns', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		for (let r = 10; r < 20; r++) {
			sheet.rows.set(r, { hidden: true });
		}
		sheet.cols = [{ min: 2, max: 3, hidden: true }];
		const layout = paginateSheet(wb, 0);
		expect(bodies(layout)).toEqual(['0-54/0-11', '55-99/0-11', '0-54/12-19', '55-99/12-19']);
	});

	it('uses row heights and column widths', () => {
		const wb = book('B10');
		const sheet = wb.sheets[0];
		sheet.rows.set(0, { height: 600, customHeight: true });
		sheet.cols = [{ min: 0, max: 0, width: 100, customWidth: true }];
		// Column A: 700 px = 525 pt gets a page of its own; row 1 leaves room for 5 more rows.
		expect(bodies(paginateSheet(wb, 0))).toEqual(['0-5/0-0', '6-9/0-0', '0-5/1-1', '6-9/1-1']);
	});

	it('returns no pages for an empty sheet', () => {
		const wb = createWorkbook();
		const layout = paginateSheet(wb, 0);
		expect(layout.pages).toEqual([]);
		expect(layout.areas).toEqual([]);
	});
});

describe('paginateSheet: manual breaks', () => {
	it('starts new pages at manual breaks', () => {
		const wb = book();
		wb.sheets[0].rowBreaks = [20];
		wb.sheets[0].colBreaks = [5];
		expect(bodies(paginateSheet(wb, 0))).toEqual([
			'0-19/0-4',
			'20-64/0-4',
			'65-99/0-4',
			'0-19/5-14',
			'20-64/5-14',
			'65-99/5-14',
			'0-19/15-19',
			'20-64/15-19',
			'65-99/15-19',
		]);
	});

	it('ignores manual breaks when fitting to pages', () => {
		const wb = book('J40');
		wb.sheets[0].rowBreaks = [20];
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
		const layout = paginateSheet(wb, 0);
		expect(layout.scale).toBe(1);
		expect(bodies(layout)).toEqual(['0-39/0-9']);
	});
});

describe('paginateSheet: scaling', () => {
	it('applies an Adjust to percentage', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { scale: 50 };
		const layout = paginateSheet(wb, 0);
		expect(layout.scale).toBe(0.5);
		expect(bodies(layout)).toEqual(['0-90/0-19', '91-99/0-19']);
		expect(layout.pages[0].blocks[0].width).toBeCloseTo(480);
	});

	it('clamps the scale to 10-400%', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { scale: 5 };
		expect(paginateSheet(wb, 0).scale).toBe(0.1);
		wb.sheets[0].pageSetup = { scale: 999 };
		expect(paginateSheet(wb, 0).scale).toBe(4);
	});

	it('fits all columns on one page', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
		const layout = paginateSheet(wb, 0);
		// 20 columns of 48 pt fit 511.2 pt at 53% (25.44 pt each), not at 54%.
		expect(layout.scale).toBe(0.53);
		expect(layout.fitToPage).toBe(true);
		expect(bodies(layout)).toEqual(['0-85/0-19', '86-99/0-19']);
	});

	it('fits all rows on one page', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 0, fitToHeight: 1 };
		const layout = paginateSheet(wb, 0);
		// 100 rows of 15 pt fit 684 pt at 45%, not at 46%.
		expect(layout.scale).toBe(0.45);
		expect(layout.pages).toHaveLength(1);
	});

	it('fits the sheet on one page with the smaller of both scales', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
		const layout = paginateSheet(wb, 0);
		expect(layout.scale).toBe(0.45);
		expect(bodies(layout)).toEqual(['0-99/0-19']);
	});

	it('defaults fitToWidth / fitToHeight to one page and never enlarges', () => {
		const wb = book('C5');
		wb.sheets[0].pageSetup = { fitToPage: true };
		expect(paginateSheet(wb, 0).scale).toBe(1);
		const big = book();
		big.sheets[0].pageSetup = { fitToPage: true };
		expect(paginateSheet(big, 0).scale).toBe(0.45);
	});

	it('fits to several pages', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 0, fitToHeight: 2 };
		const layout = paginateSheet(wb, 0);
		// 100 rows on two pages: 50 rows of 13.65 pt each at 91%.
		expect(layout.scale).toBe(0.91);
		expect(new Set(layout.pages.map((p) => p.body.top)).size).toBe(2);
	});

	it('stops at 10% when nothing fits', () => {
		const wb = book('A5000');
		wb.sheets[0].pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
		expect(paginateSheet(wb, 0).scale).toBe(0.1);
	});

	it('ignores the scale while fitting', () => {
		const wb = book('C5');
		wb.sheets[0].pageSetup = { fitToPage: true, scale: 30 };
		expect(paginateSheet(wb, 0).scale).toBe(1);
	});
});

describe('paginateSheet: print area and titles', () => {
	it('prints the print area, each area on its own pages', () => {
		const wb = book();
		wb.definedNames.push({
			name: '_xlnm.Print_Area',
			formula: 'Sheet1!$A$1:$B$5,Sheet1!$D$10:$E$12',
			localSheet: 0,
			builtin: true,
		});
		const layout = paginateSheet(wb, 0);
		expect(layout.areas).toEqual([range('A1:B5'), range('D10:E12')]);
		expect(bodies(layout)).toEqual(['0-4/0-1', '9-11/3-4']);
		expect(layout.pages.map((p) => p.area)).toEqual([0, 1]);
		expect(bodies(paginateSheet(wb, 0, { ignorePrintArea: true }))).toHaveLength(6);
	});

	it('clamps whole-column print areas to the used rows', () => {
		const wb = book('C30');
		wb.definedNames.push({
			name: '_xlnm.Print_Area',
			formula: 'Sheet1!$B:$C',
			localSheet: 0,
			builtin: true,
		});
		expect(paginateSheet(wb, 0).areas).toEqual([range('B1:C30')]);
	});

	it('prints a requested range or ranges instead (Print Selection)', () => {
		const wb = book();
		expect(bodies(paginateSheet(wb, 0, { range: range('C3:D4') }))).toEqual(['2-3/2-3']);
		const multi = paginateSheet(wb, 0, { ranges: [range('A1'), range('C3')] });
		expect(bodies(multi)).toEqual(['0-0/0-0', '2-2/2-2']);
	});

	it('repeats title rows and columns on the following pages', () => {
		const wb = book();
		wb.definedNames.push({
			name: '_xlnm.Print_Titles',
			formula: 'Sheet1!$A:$A,Sheet1!$1:$1',
			localSheet: 0,
			builtin: true,
		});
		const layout = paginateSheet(wb, 0);
		expect(layout.titleRows).toEqual({ start: 0, end: 0 });
		expect(layout.titleCols).toEqual({ start: 0, end: 0 });
		const [p1, p2] = layout.pages;
		expect(p1.titleRows).toBeUndefined();
		expect(p1.body).toEqual(range('A1:J45'));
		expect(p1.blocks.map((b) => b.kind)).toEqual(['body']);
		// Page 2 repeats row 1 above rows 46+: one row less room.
		expect(p2.titleRows).toEqual({ start: 0, end: 0 });
		expect(p2.body).toEqual(range('A46:J89'));
		expect(p2.blocks.map((b) => b.kind)).toEqual(['titleRows', 'body']);
		expect(p2.blocks[0].y).toBe(54);
		expect(p2.blocks[1].y).toBe(69);
		const p4 = layout.pages[3];
		expect(p4.titleCols).toEqual({ start: 0, end: 0 });
		expect(p4.titleRows).toBeUndefined();
		expect(p4.body).toEqual(range('K1:S45'));
		const p5 = layout.pages[4];
		expect(p5.blocks.map((b) => b.kind)).toEqual(['corner', 'titleRows', 'titleCols', 'body']);
		const [corner, rows, cols, body] = p5.blocks;
		expect(corner.range).toEqual(range('A1'));
		expect(rows.range).toEqual(range('K1:S1'));
		expect(cols.range).toEqual(range('A46:A89'));
		expect(body.range).toEqual(range('K46:S89'));
		expect(body.x).toBeCloseTo(50.4 + 48);
		expect(body.y).toBe(54 + 15);
	});

	it('repeats titles above a print area that starts below them', () => {
		const wb = book();
		wb.definedNames.push(
			{ name: '_xlnm.Print_Area', formula: 'Sheet1!$A$10:$C$12', localSheet: 0, builtin: true },
			{ name: '_xlnm.Print_Titles', formula: 'Sheet1!$1:$2', localSheet: 0, builtin: true },
		);
		const page = paginateSheet(wb, 0).pages[0];
		expect(page.titleRows).toEqual({ start: 0, end: 1 });
		expect(page.blocks.map((b) => b.kind)).toEqual(['titleRows', 'body']);
		expect(page.content.height).toBe(75);
	});
});

describe('paginateSheet: options', () => {
	it('centers the content horizontally and vertically', () => {
		const wb = book('B2');
		wb.sheets[0].printOptions = { horizontalCentered: true, verticalCentered: true };
		const page = paginateSheet(wb, 0).pages[0];
		expect(page.content.x).toBeCloseTo(50.4 + (511.2 - 96) / 2);
		expect(page.content.y).toBe(54 + (684 - 30) / 2);
		expect(page.blocks[0].x).toBeCloseTo(page.content.x);
	});

	it('reserves room for printed headings', () => {
		const wb = book();
		wb.sheets[0].printOptions = { headings: true, gridLines: true };
		const layout = paginateSheet(wb, 0);
		expect(layout.headings).toBe(true);
		expect(layout.gridLines).toBe(true);
		const page = layout.pages[0];
		expect(page.headings).toEqual({ rowWidth: 23.25, colHeight: 15 });
		expect(page.blocks[0].x).toBeCloseTo(50.4 + 23.25);
		expect(page.blocks[0].y).toBe(54 + 15);
		// 488 pt for columns (10 still fit), 669 pt for rows (44).
		expect(page.body).toEqual(range('A1:J44'));
	});

	it('numbers pages from the first page number', () => {
		const wb = book();
		wb.sheets[0].pageSetup = { firstPageNumber: 5, useFirstPageNumber: true };
		expect(paginateSheet(wb, 0).pages.map((p) => p.pageNumber)).toEqual([5, 6, 7, 8, 9, 10]);
		wb.sheets[0].pageSetup = { firstPageNumber: 5 };
		expect(paginateSheet(wb, 0).pages[0].pageNumber).toBe(1);
		expect(paginateSheet(wb, 0, { firstPageNumber: 4 }).pages[0].pageNumber).toBe(4);
	});

	it('formats headers and footers per page', () => {
		const wb = book();
		wb.sheets[0].headerFooter = {
			oddHeader: '&L&A&RPage &P of &N',
			oddFooter: '&C&F',
			firstHeader: 'Cover',
			differentFirst: true,
		};
		const layout = paginateSheet(wb, 0, { fileName: 'Book.xlsx', totalPages: 9 });
		const [p1, p2] = layout.pages;
		expect(p1.header?.center[0].text).toBe('Cover');
		expect(p1.footer).toBeNull();
		expect(p2.header?.left[0].text).toBe('Sheet1');
		expect(p2.header?.right[0].text).toBe('Page 2 of 9');
		expect(p2.footer?.center[0].text).toBe('Book.xlsx');
		expect(layout.headerFooterScale).toBe(1);
	});

	it('includes floating objects in the printed range', () => {
		const wb = book('B2');
		wb.sheets[0].drawings.push({
			id: '1',
			kind: 'shape',
			anchor: {
				type: 'twoCell',
				from: { row: 3, col: 4, rowOff: 0, colOff: 0 },
				to: { row: 8, col: 6, rowOff: 0, colOff: 0 },
			},
		} as Workbook['sheets'][number]['drawings'][number]);
		expect(paginateSheet(wb, 0).areas).toEqual([range('A1:F8')]);
	});
});

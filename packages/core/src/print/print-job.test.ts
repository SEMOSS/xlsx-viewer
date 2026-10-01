import { describe, expect, it } from 'vitest';

import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import { paginateSheets, totalPrintPages } from './print-job';

describe('paginateSheets', () => {
	it('continues page numbers across sheets and counts the whole job', () => {
		const wb = createWorkbook({ sheetNames: ['A', 'B', 'C'] });
		putCell(wb.sheets[0], 0, 0, { value: 1 });
		putCell(wb.sheets[0], 99, 0, { value: 1 });
		putCell(wb.sheets[1], 0, 0, { value: 1 });
		putCell(wb.sheets[2], 0, 0, { value: 1 });
		wb.sheets[2].pageSetup = { firstPageNumber: 10, useFirstPageNumber: true };
		for (const s of wb.sheets) {
			s.headerFooter = { oddFooter: 'Page &P of &N' };
		}
		const layouts = paginateSheets(wb, [0, 1, 2]);
		expect(layouts.map((l) => l.pages.map((p) => p.pageNumber))).toEqual([[1, 2, 3], [4], [10]]);
		expect(totalPrintPages(layouts)).toBe(5);
		expect(layouts.map((l) => l.totalPages)).toEqual([5, 5, 5]);
		expect(layouts[1].pages[0].footer?.center[0].text).toBe('Page 4 of 5');
		expect(layouts[2].pages[0].footer?.center[0].text).toBe('Page 10 of 5');
	});

	it('numbers a single sheet on its own', () => {
		const wb = createWorkbook();
		putCell(wb.sheets[0], 0, 0, { value: 1 });
		wb.sheets[0].headerFooter = { oddHeader: '&P/&N' };
		const [layout] = paginateSheets(wb, [0]);
		expect(layout.pages[0].header?.center[0].text).toBe('1/1');
		expect(paginateSheets(wb, [])).toEqual([]);
	});
});

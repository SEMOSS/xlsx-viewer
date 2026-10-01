import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, paginateSheets, putCell } from 'xlsx-viewer-core';

import { pageName, printCss, printWorkbook } from './print-document';

function book() {
	const wb = createWorkbook({ sheetNames: ['One', 'Two'] });
	putCell(wb.sheets[0], 0, 0, { value: 'First' });
	putCell(wb.sheets[0], 59, 0, { value: 'Last' });
	putCell(wb.sheets[1], 0, 0, { value: 'Other' });
	wb.sheets[1].pageSetup = { orientation: 'landscape', paperSize: 9 };
	wb.sheets[0].headerFooter = { oddFooter: '&CPage &P of &N' };
	return wb;
}

describe('print document', () => {
	it('sizes pages with @page rules, named per paper size', () => {
		const layouts = paginateSheets(book(), [0, 1]);
		const css = printCss(layouts);
		expect(pageName(layouts[0])).toBe('xvp612x792');
		expect(css).toContain('@page { size: 612pt 792pt; margin: 0; }');
		expect(css).toContain(`@page ${pageName(layouts[1])} { size: ${layouts[1].paper.width}pt`);
		expect(css).toContain('print-color-adjust: exact');
		expect(css).toContain('break-after: page');
	});

	it('renders every page into a hidden iframe, prints and cleans up', async () => {
		const wb = book();
		let pages: Element[] = [];
		let footer = '';
		const print = vi.fn((win: Window) => {
			pages = Array.from(win.document.querySelectorAll('.xv-print-sheet'));
			footer = win.document.querySelector('.xv-print-footer')?.textContent ?? '';
			win.dispatchEvent(new Event('afterprint'));
		});
		const count = await printWorkbook(wb, { sheets: [0, 1], print });
		expect(count).toBe(3);
		expect(print).toHaveBeenCalledTimes(1);
		expect(pages).toHaveLength(3);
		expect(pages[2].className).toContain(pageName(paginateSheets(wb, [1])[0]));
		expect(footer).toBe('Page 1 of 3');
		await new Promise((r) => setTimeout(r, 10));
		expect(document.querySelectorAll('iframe')).toHaveLength(0);
	});

	it('prints nothing for an empty sheet', async () => {
		const print = vi.fn();
		expect(await printWorkbook(createWorkbook(), { print })).toBe(0);
		expect(print).not.toHaveBeenCalled();
	});
});

import { describe, expect, it } from 'vitest';
import type { DrawingObject, Workbook } from 'xlsx-viewer-core';
import { createWorkbook, paginateSheet, putCell } from 'xlsx-viewer-core';

import { buildPrintPageModel } from './page-model';
import { createPrintGeometry, PX_PER_PT } from './print-geometry';

function book(): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	putCell(sheet, 0, 0, { value: 'Title' });
	putCell(sheet, 99, 14, { value: 'End' });
	return wb;
}

function shape(id: string, row: number, col: number): DrawingObject {
	return {
		id,
		kind: 'shape',
		anchor: {
			type: 'twoCell',
			from: { row, col, rowOff: 0, colOff: 0 },
			to: { row: row + 2, col: col + 2, rowOff: 0, colOff: 0 },
		},
	} as DrawingObject;
}

describe('createPrintGeometry', () => {
	it('matches the printed sizes at the print scale without rounding', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		sheet.rows.set(3, { height: 14.4 });
		sheet.rows.set(4, { hidden: true });
		sheet.cols = [{ min: 1, max: 1, width: 20 }];
		const g = createPrintGeometry(sheet, 0.5);
		expect(g.zoom).toBe(0.5);
		expect(g.rowHeight(0)).toBe(10);
		expect(g.rowHeight(3)).toBeCloseTo(9.6);
		expect(g.rowHeight(4)).toBe(0);
		expect(g.colWidth(0)).toBe(32);
		expect(g.colWidth(1)).toBe(70);
		expect(g.rowTop(5)).toBeCloseTo(10 * 3 + 9.6);
		expect(g.rangeRect({ top: 0, left: 0, bottom: 1, right: 1 })).toEqual({
			x: 0,
			y: 0,
			width: 102,
			height: 20,
		});
		expect(g.cellAt(40, 25)).toEqual({ row: 2, col: 1 });
	});
});

describe('buildPrintPageModel', () => {
	it('places blocks where the paginator put them, in 96 DPI pixels', () => {
		const wb = book();
		const layout = paginateSheet(wb, 0);
		const sheet = wb.sheets[0];
		const g = createPrintGeometry(sheet, layout.scale);
		const page = layout.pages[1];
		const model = buildPrintPageModel(layout, page, sheet, g);
		expect(model.width).toBe(816);
		expect(model.height).toBe(1056);
		const [body] = model.blocks;
		expect(body.kind).toBe('body');
		expect(body.x).toBeCloseTo(0.7 * 96);
		expect(body.y).toBe(72);
		expect(body.rows).toEqual({ start: 45, end: 89 });
		expect(body.originY).toBe(g.rowTop(45));
		// The block is exactly as tall as its rows in the print geometry.
		expect(body.height).toBeCloseTo(g.rowTop(90) - g.rowTop(45));
		expect(model.headings).toEqual([]);
		expect(model.header).toBeNull();
	});

	it('lays out headings, title blocks and header / footer bands', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		sheet.printOptions = { headings: true };
		sheet.headerFooter = { oddHeader: '&LLeft&RPage &P', oddFooter: '&C&A' };
		wb.definedNames.push({
			name: '_xlnm.Print_Titles',
			formula: 'Sheet1!$1:$1',
			localSheet: 0,
			builtin: true,
		});
		const layout = paginateSheet(wb, 0);
		const page = layout.pages[1];
		const model = buildPrintPageModel(layout, page, sheet, createPrintGeometry(sheet, 1));
		expect(model.blocks.map((b) => b.kind)).toEqual(['titleRows', 'body']);
		const corner = model.headings.find((h) => h.key === 'corner')!;
		expect(corner.width).toBeCloseTo(23.25 * PX_PER_PT);
		expect(corner.height).toBe(20);
		const rowLabels = model.headings.filter((h) => h.kind === 'row').map((h) => h.text);
		expect(rowLabels[0]).toBe('1');
		expect(rowLabels[1]).toBe(String(page.body.top + 1));
		const colLabels = model.headings.filter((h) => h.kind === 'col').map((h) => h.text);
		expect(colLabels.slice(0, 2)).toEqual(['A', 'B']);
		expect(model.header?.offset).toBeCloseTo(0.3 * 96);
		expect(model.header?.left).toBeCloseTo(0.7 * 96);
		expect(model.header?.right).toBeCloseTo((8.5 - 0.7) * 96);
		expect(model.header?.sections.right[0].text).toBe('Page 2');
		expect(model.footer?.sections.center[0].text).toBe('Sheet1');
	});

	it('keeps the floating objects that reach into the body', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		sheet.drawings = [
			shape('in', 2, 2),
			shape('below', 60, 2),
			{ ...shape('hidden', 3, 3), hidden: true },
		];
		const layout = paginateSheet(wb, 0);
		const g = createPrintGeometry(sheet, 1);
		const first = buildPrintPageModel(layout, layout.pages[0], sheet, g);
		expect(first.drawings.map((d) => d.object.id)).toEqual(['in']);
		expect(first.drawings[0].rect).toEqual({ x: 128, y: 40, width: 128, height: 40 });
		const second = buildPrintPageModel(layout, layout.pages[1], sheet, g);
		expect(second.drawings.map((d) => d.object.id)).toEqual(['below']);
		sheet.pageSetup = { draft: true };
		const draft = paginateSheet(wb, 0);
		expect(buildPrintPageModel(draft, draft.pages[0], sheet, g).drawings).toEqual([]);
	});
});

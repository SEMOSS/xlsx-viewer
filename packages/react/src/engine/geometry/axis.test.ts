import { describe, expect, it } from 'vitest';

import { createWorksheet } from 'xlsx-viewer-core';

import { AxisMetrics } from './axis';
import { SheetGeometry } from './sheet-geometry';

describe('AxisMetrics', () => {
	const axis = new AxisMetrics(
		20,
		new Map([
			[2, 40],
			[5, 0],
			[6, 0],
			[9, 10],
		]),
		100,
	);

	it('computes offsets with overrides', () => {
		expect(axis.offset(0)).toBe(0);
		expect(axis.offset(2)).toBe(40);
		expect(axis.offset(3)).toBe(80);
		expect(axis.offset(5)).toBe(120);
		expect(axis.offset(7)).toBe(120);
		expect(axis.offset(10)).toBe(170);
		expect(axis.total()).toBe(100 * 20 + 20 - 40 - 10);
	});

	it('hit-tests positions, skipping hidden lines', () => {
		expect(axis.indexAt(0)).toBe(0);
		expect(axis.indexAt(39)).toBe(1);
		expect(axis.indexAt(40)).toBe(2);
		expect(axis.indexAt(79)).toBe(2);
		expect(axis.indexAt(119)).toBe(4);
		expect(axis.indexAt(120)).toBe(7);
		expect(axis.indexAt(1e9)).toBe(99);
	});
});

describe('SheetGeometry', () => {
	it('uses Excel default sizes (64 x 20 px)', () => {
		const g = new SheetGeometry(createWorksheet('S', 1), { zoom: 1 });
		expect(g.colWidth(0)).toBe(64);
		expect(g.rowHeight(0)).toBe(20);
		expect(g.cellAt(130, 45)).toEqual({ row: 2, col: 2 });
	});

	it('applies custom widths, hidden columns and zoom', () => {
		const sheet = createWorksheet('S', 1);
		sheet.cols = [
			{ min: 1, max: 1, width: 20, customWidth: true },
			{ min: 2, max: 2, hidden: true },
		];
		sheet.rows.set(0, { height: 30, customHeight: true });
		const g = new SheetGeometry(sheet, { zoom: 2 });
		expect(g.colWidth(1)).toBe(140 * 2);
		expect(g.colWidth(2)).toBe(0);
		expect(g.rowHeight(0)).toBe(80);
		expect(g.rangeRect({ top: 0, left: 0, bottom: 1, right: 3 })).toEqual({
			x: 0,
			y: 0,
			width: 128 + 280 + 0 + 128,
			height: 80 + 40,
		});
	});
});

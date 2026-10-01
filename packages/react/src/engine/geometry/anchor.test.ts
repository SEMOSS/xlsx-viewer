import { describe, expect, it } from 'vitest';
import type { DrawingAnchor } from 'xlsx-viewer-core';
import { createWorksheet, EMU_PER_PIXEL } from 'xlsx-viewer-core';

import { anchorRect, rectToAnchor } from './anchor';
import { SheetGeometry } from './sheet-geometry';

const px = (n: number) => n * EMU_PER_PIXEL;

function geometry(zoom = 1): SheetGeometry {
	const sheet = createWorksheet('S', 1);
	sheet.cols = [
		{ min: 1, max: 1, width: 20 },
		{ min: 4, max: 4, hidden: true },
	];
	sheet.rows.set(3, { height: 30 });
	return new SheetGeometry(sheet, { zoom });
}

// Columns: A 0-64, B 64-204 (140 px), C 204-268, D 268-332, E hidden, F 332-396.
// Rows: 1-3 are 20 px, row 4 is 40 px (y 60-100), row 5 starts at 100.
const g = geometry();

const twoCell: DrawingAnchor = {
	type: 'twoCell',
	editAs: 'oneCell',
	from: { col: 1, colOff: px(10), row: 2, rowOff: px(5) },
	to: { col: 3, colOff: 0, row: 4, rowOff: px(10) },
};

describe('anchorRect', () => {
	it('places two-cell anchors between their corner cells plus offsets', () => {
		expect(g.colLeft(1)).toBe(64);
		expect(anchorRect(twoCell, g)).toEqual({ x: 74, y: 45, width: 268 - 74, height: 110 - 45 });
	});

	it('clamps offsets to the size of their cell', () => {
		const anchor: DrawingAnchor = {
			type: 'twoCell',
			from: { col: 0, colOff: px(500), row: 0, rowOff: px(50) },
			to: { col: 2, colOff: px(10), row: 1, rowOff: px(10) },
		};
		expect(anchorRect(anchor, g)).toEqual({ x: 64, y: 20, width: 214 - 64, height: 10 });
	});

	it('collapses offsets inside hidden columns and never returns a negative size', () => {
		const hidden: DrawingAnchor = {
			type: 'twoCell',
			from: { col: 4, colOff: px(30), row: 0, rowOff: 0 },
			to: { col: 5, colOff: px(4), row: 1, rowOff: 0 },
		};
		expect(anchorRect(hidden, g)).toEqual({ x: 332, y: 0, width: 4, height: 20 });
		const backwards: DrawingAnchor = {
			type: 'twoCell',
			from: { col: 3, colOff: 0, row: 4, rowOff: 0 },
			to: { col: 1, colOff: 0, row: 2, rowOff: 0 },
		};
		expect(anchorRect(backwards, g)).toMatchObject({ width: 0, height: 0 });
	});

	it('sizes one-cell anchors by their extent and absolute anchors by position and extent', () => {
		const oneCell: DrawingAnchor = {
			type: 'oneCell',
			from: { col: 2, colOff: px(4), row: 1, rowOff: px(2) },
			ext: { cx: px(100), cy: px(50) },
		};
		expect(anchorRect(oneCell, g)).toEqual({ x: 208, y: 22, width: 100, height: 50 });
		const absolute: DrawingAnchor = {
			type: 'absolute',
			pos: { x: px(15), y: px(25) },
			ext: { cx: px(30), cy: px(40) },
		};
		expect(anchorRect(absolute, g)).toEqual({ x: 15, y: 25, width: 30, height: 40 });
	});

	it('scales with zoom', () => {
		const z = geometry(2);
		expect(anchorRect(twoCell, z)).toEqual({ x: 148, y: 90, width: 536 - 148, height: 220 - 90 });
		const absolute: DrawingAnchor = {
			type: 'absolute',
			pos: { x: px(15), y: px(25) },
			ext: { cx: px(30), cy: px(40) },
		};
		expect(anchorRect(absolute, z)).toEqual({ x: 30, y: 50, width: 60, height: 80 });
	});
});

describe('rectToAnchor', () => {
	it('keeps the anchor type and two-cell editAs', () => {
		const next = rectToAnchor({ x: 74, y: 45, width: 194, height: 65 }, twoCell, g);
		expect(next).toEqual(twoCell);
	});

	it('round-trips two-cell anchors at any zoom', () => {
		for (const zoom of [0.5, 1, 1.5, 2]) {
			const z = geometry(zoom);
			expect(rectToAnchor(anchorRect(twoCell, z), twoCell, z), `zoom ${zoom}`).toEqual(twoCell);
		}
	});

	it('places a corner on a cell boundary at the start of the next cell', () => {
		const next = rectToAnchor({ x: 64, y: 60, width: 140, height: 40 }, twoCell, g);
		expect(next).toEqual({
			type: 'twoCell',
			editAs: 'oneCell',
			from: { col: 1, colOff: 0, row: 3, rowOff: 0 },
			to: { col: 2, colOff: 0, row: 4, rowOff: 0 },
		});
	});

	it('skips hidden columns when a corner lands on them', () => {
		const next = rectToAnchor({ x: 332, y: 0, width: 10, height: 10 }, twoCell, g);
		expect(next.type === 'twoCell' && next.from).toEqual({ col: 5, colOff: 0, row: 0, rowOff: 0 });
	});

	it('stores extents in unzoomed EMU for one-cell anchors', () => {
		const prev: DrawingAnchor = {
			type: 'oneCell',
			from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
			ext: { cx: 1, cy: 1 },
		};
		const z = geometry(2);
		const next = rectToAnchor({ x: 148, y: 90, width: 200, height: 100 }, prev, z);
		expect(next).toEqual({
			type: 'oneCell',
			from: { col: 1, colOff: px(10), row: 2, rowOff: px(5) },
			ext: { cx: px(100), cy: px(50) },
		});
		expect(anchorRect(next, z)).toEqual({ x: 148, y: 90, width: 200, height: 100 });
	});

	it('stores position and extent for absolute anchors', () => {
		const prev: DrawingAnchor = { type: 'absolute', pos: { x: 0, y: 0 }, ext: { cx: 0, cy: 0 } };
		const z = geometry(1.5);
		const rect = { x: 30, y: 45, width: 90, height: 60 };
		const next = rectToAnchor(rect, prev, z);
		expect(next).toEqual({
			type: 'absolute',
			pos: { x: px(20), y: px(30) },
			ext: { cx: px(60), cy: px(40) },
		});
		expect(anchorRect(next, z)).toEqual(rect);
	});

	it('round-trips arbitrary rectangles through each anchor type', () => {
		const types: DrawingAnchor[] = [
			twoCell,
			{ type: 'oneCell', from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, ext: { cx: 0, cy: 0 } },
			{ type: 'absolute', pos: { x: 0, y: 0 }, ext: { cx: 0, cy: 0 } },
		];
		const rects = [
			{ x: 0, y: 0, width: 10, height: 10 },
			{ x: 100, y: 75, width: 333, height: 17 },
			{ x: 401, y: 999, width: 1, height: 250 },
		];
		for (const prev of types) {
			for (const rect of rects) {
				const back = anchorRect(rectToAnchor(rect, prev, g), g);
				expect(back, `${prev.type} ${JSON.stringify(rect)}`).toEqual(rect);
			}
		}
	});
});

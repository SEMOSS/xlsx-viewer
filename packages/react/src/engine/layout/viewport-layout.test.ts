import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { BorderEdge, BorderStyle, FillSpec, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	applyBorderPreset,
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	mergeRange,
	parseRange,
	setCellInput,
	setColumnsHidden,
	setRowsHidden,
	StyleResolver,
} from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import type { DescribeOptions } from '../render/cell-content';
import { borderStroke } from '../render/borders';
import { setMeasureContextForTests } from '../render/text-measure';
import { MergeIndex } from './style-at';
import type { ViewportLayout } from './types';
import { layoutViewport, type LayoutParams } from './viewport-layout';

/** Fake canvas: glyphs are 10 px wide in the default 11 pt font and scale with the font size. */
const fakeCanvas = {
	font: '',
	measureText(text: string) {
		const size = Number(/([\d.]+)px/.exec(fakeCanvas.font)?.[1] ?? 0);
		return { width: text.length * Math.round((size * 15) / 22) };
	},
};
beforeAll(() => setMeasureContextForTests(fakeCanvas as unknown as CanvasRenderingContext2D));
afterAll(() => setMeasureContextForTests(null));

interface Setup {
	/** Text typed per cell, `{ A1: 'abc' }`. */
	cells?: Record<string, string>;
	/** Style patches per range, applied after typing. */
	styles?: Record<string, StylePatch>;
	build?: (tx: EditTransaction, wb: Workbook) => void;
}

function book({ cells = {}, styles = {}, build }: Setup = {}): Workbook {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	for (const [ref, text] of Object.entries(cells)) {
		const { top, left } = parseRange(ref)!;
		setCellInput(tx, wb, 0, top, left, text);
	}
	for (const [ref, patch] of Object.entries(styles)) {
		applyStylePatch(tx, wb, 0, [parseRange(ref)!], patch);
	}
	build?.(tx, wb);
	tx.commit();
	return wb;
}

type Options = Partial<Pick<LayoutParams, 'rows' | 'cols' | 'canvasColor'>> & {
	zoom?: number;
	describe?: Partial<DescribeOptions>;
};

/** Lays out rows 1-10 and columns A-H at 100% unless told otherwise. */
function layout(
	wb: Workbook,
	{ zoom = 1, describe: extra, ...rest }: Options = {},
): ViewportLayout {
	const sheet = wb.sheets[0];
	return layoutViewport({
		rows: { start: 0, end: 9 },
		cols: { start: 0, end: 7 },
		...rest,
		sheet,
		geometry: new SheetGeometry(sheet, { zoom }),
		styles: new StyleResolver(wb.styles, wb.theme),
		merges: new MergeIndex(sheet.merges),
		describe: {
			zoom,
			colorCtx: { theme: wb.theme },
			date1904: false,
			defaultTextColor: '#000000',
			...extra,
		},
	});
}

const solid = (rgb: string): FillSpec => ({
	type: 'pattern',
	patternType: 'solid',
	fgColor: { rgb },
});
const edge = (style: BorderStyle, rgb?: string): BorderEdge => ({
	style,
	color: rgb ? { rgb } : undefined,
});

const NAMES: BorderStyle[] = ['hair', 'dotted', 'dashed', 'thin', 'medium', 'thick', 'double'];

/** Border segments as sorted `h20 64-128 thin #000000` strings (line, span, style, color). */
function edges(l: ViewportLayout): string[] {
	return l.borders
		.map((s) => {
			const name = NAMES.find((n) => borderStroke(n) === s.stroke) ?? '?';
			return `${s.orientation}${s.at} ${s.from}-${s.to} ${name} ${s.color}`;
		})
		.sort();
}

describe('layoutViewport: fills', () => {
	it('paints cell fills over the cell rectangle in sheet pixels', () => {
		const pattern: FillSpec = {
			type: 'pattern',
			patternType: 'lightGray',
			fgColor: { rgb: 'FFFF0000' },
			bgColor: { rgb: 'FF00FF00' },
		};
		const l = layout(book({ styles: { B2: { fill: solid('FFFFFF00') }, D3: { fill: pattern } } }));
		expect(l.fills.map((f) => f.key)).toEqual(['f:1:1', 'f:2:3']);
		expect(l.fills[0]).toEqual({
			key: 'f:1:1',
			x: 64,
			y: 20,
			width: 64,
			height: 20,
			css: { backgroundColor: '#FFFF00' },
		});
		expect(l.fills[1]).toMatchObject({ x: 192, y: 40, css: { backgroundColor: '#00FF00' } });
		expect(l.fills[1].css.backgroundImage).toMatch(/^url\("data:image\/svg\+xml,/);
	});

	it('paints column and row formats on empty cells', () => {
		const column = layout(book({ styles: { 'C:C': { fill: solid('FF0070C0') } } }));
		expect(column.fills).toHaveLength(10);
		expect(column.fills.every((f) => f.x === 128 && f.css.backgroundColor === '#0070C0')).toBe(
			true,
		);
		const row = layout(book({ styles: { '3:3': { fill: solid('FF0070C0') } } }));
		expect(row.fills.map((f) => f.key)).toEqual([0, 1, 2, 3, 4, 5, 6, 7].map((c) => `f:2:${c}`));
	});

	it('only lays out the visible rows and columns', () => {
		const wb = book({
			styles: { A1: { fill: solid('FFFF0000') }, E6: { fill: solid('FF00FF00') } },
		});
		const l = layout(wb, { rows: { start: 5, end: 6 }, cols: { start: 3, end: 4 } });
		expect(l.fills).toEqual([
			{ key: 'f:5:4', x: 256, y: 100, width: 64, height: 20, css: { backgroundColor: '#00FF00' } },
		]);
	});

	it('scales boxes with zoom', () => {
		const l = layout(book({ styles: { B2: { fill: solid('FFFFFF00') } } }), { zoom: 2 });
		expect(l.fills[0]).toMatchObject({ x: 128, y: 40, width: 128, height: 40 });
	});
});

describe('layoutViewport: borders', () => {
	it('draws the four edges of a bordered cell on its gridlines', () => {
		const l = layout(
			book({ build: (tx, wb) => applyBorderPreset(tx, wb, 0, parseRange('B2')!, 'all') }),
		);
		expect(edges(l)).toEqual(
			[
				'h20 64-128 thin #000000',
				'h40 64-128 thin #000000',
				'v64 20-40 thin #000000',
				'v128 20-40 thin #000000',
			].sort(),
		);
		expect(l.borders.every((b) => b.zoom === 1)).toBe(true);
	});

	it('draws shared edges once', () => {
		const l = layout(
			book({ build: (tx, wb) => applyBorderPreset(tx, wb, 0, parseRange('B2:C3')!, 'all') }),
		);
		// 3 horizontal lines x 2 columns + 3 vertical lines x 2 rows.
		expect(l.borders).toHaveLength(12);
	});

	it('lets the stronger style win a shared edge, whichever cell draws it', () => {
		const wb = book({
			styles: {
				B2: { border: { bottom: edge('thick', 'FFFF0000') } },
				B3: { border: { top: edge('thin', 'FF0000FF') } },
				B5: { border: { bottom: edge('thin', 'FFFF0000') } },
				B6: { border: { top: edge('medium', 'FF0000FF') } },
				E2: { border: { right: edge('double', 'FFFF0000') } },
				F2: { border: { left: edge('thick', 'FF0000FF') } },
			},
		});
		expect(edges(layout(wb))).toEqual(
			[
				'h40 64-128 thick #FF0000',
				'h100 64-128 medium #0000FF',
				'v320 20-40 double #FF0000',
			].sort(),
		);
	});

	it('gives equal styles to the right or lower cell', () => {
		const wb = book({
			styles: {
				B2: { border: { right: edge('thin', 'FFFF0000'), bottom: edge('dashed', 'FFFF0000') } },
				C2: { border: { left: edge('thin', 'FF0000FF') } },
				B3: { border: { top: edge('dashed', 'FF00FF00') } },
			},
		});
		expect(edges(layout(wb))).toEqual(
			['h40 64-128 dashed #00FF00', 'v128 20-40 thin #0000FF'].sort(),
		);
	});

	it('resolves theme and automatic border colors', () => {
		const wb = book({
			styles: {
				A1: { border: { top: { style: 'thin', color: { theme: 4 } }, left: edge('thin') } },
			},
		});
		expect(edges(layout(wb))).toEqual(['h0 0-64 thin #4472C4', 'v0 0-20 thin #000000'].sort());
	});
});

describe('layoutViewport: merged cells', () => {
	const merged = (setup: Setup = {}) =>
		book({
			...setup,
			build: (tx, wb) => {
				setup.build?.(tx, wb);
				mergeRange(tx, wb, 0, parseRange('B2:C3')!);
			},
		});

	it('lays out a merged area as one box with one text', () => {
		const l = layout(merged({ cells: { B2: 'Title' } }));
		const rect = { x: 64, y: 20, width: 128, height: 40 };
		expect(l.texts).toHaveLength(1);
		expect(l.texts[0]).toMatchObject({ key: 't:1:1', row: 1, col: 1, cell: rect, clip: rect });
		expect(l.texts[0].content.text).toBe('Title');
		expect(l.fills).toEqual([{ key: 'm:1:1', ...rect, css: { backgroundColor: '#ffffff' } }]);
		// Gridlines are hidden inside the merged area, not on its right and bottom edges.
		expect(l.gridlineMasks).toEqual([{ x: 64, y: 20, width: 127, height: 39 }]);
	});

	it('fills the area with the anchor fill, or the canvas color without one', () => {
		expect(layout(merged(), { canvasColor: '#202124' }).fills).toEqual([
			{ key: 'm:1:1', x: 64, y: 20, width: 128, height: 40, css: { backgroundColor: '#202124' } },
		]);
		const filled = layout(merged({ styles: { B2: { fill: solid('FFFFC000') } } }), {
			canvasColor: '#202124',
		});
		expect(filled.fills.map((f) => f.css.backgroundColor)).toEqual(['#FFC000']);
	});

	it('draws only the perimeter borders of the merged cells', () => {
		const l = layout(
			merged({ build: (tx, wb) => applyBorderPreset(tx, wb, 0, parseRange('B2:C3')!, 'all') }),
		);
		expect(edges(l)).toEqual(
			[
				'h20 64-128 thin #000000',
				'h20 128-192 thin #000000',
				'h60 64-128 thin #000000',
				'h60 128-192 thin #000000',
				'v64 20-40 thin #000000',
				'v64 40-60 thin #000000',
				'v192 20-40 thin #000000',
				'v192 40-60 thin #000000',
			].sort(),
		);
	});

	it('keeps the whole area when the viewport starts inside it', () => {
		const wb = book({
			cells: { A1: 'Big' },
			build: (tx, w) => mergeRange(tx, w, 0, parseRange('A1:D5')!),
		});
		const l = layout(wb, { rows: { start: 2, end: 9 }, cols: { start: 1, end: 7 } });
		expect(l.texts.map((t) => [t.key, t.cell])).toEqual([
			['t:0:0', { x: 0, y: 0, width: 256, height: 100 }],
		]);
		expect(l.fills.map((f) => f.key)).toEqual(['m:0:0']);
	});

	it('fits numbers to the merged width', () => {
		const cells = { B2: '1234567' };
		const styles = { B2: { numFmt: '#,##0' } };
		const single = layout(book({ cells, styles }));
		expect(single.texts[0]).toMatchObject({ hashes: true });
		const wide = layout(merged({ cells, styles }));
		expect(wide.texts[0]).toMatchObject({ hashes: false, content: { text: '1,234,567' } });
	});

	it('stops text overflowing from a neighbour', () => {
		const l = layout(merged({ cells: { A2: 'overflowing text' } }));
		const a2 = l.texts.find((t) => t.key === 't:1:0');
		expect(a2?.clip).toEqual(a2?.cell);
	});
});

describe('layoutViewport: hidden rows and columns', () => {
	it('skips hidden rows: no fills, borders or texts, and the next row moves up', () => {
		const wb = book({
			cells: { A1: 'a', A2: 'b', A3: 'c' },
			styles: { A2: { fill: solid('FFFF0000') } },
			build: (tx, w) => {
				applyBorderPreset(tx, w, 0, parseRange('A2')!, 'all');
				setRowsHidden(tx, w, 0, 1, 1, true);
			},
		});
		const l = layout(wb);
		expect(l.texts.map((t) => [t.key, t.cell.y])).toEqual([
			['t:0:0', 0],
			['t:2:0', 20],
		]);
		expect(l.fills).toEqual([]);
		expect(l.borders).toEqual([]);
	});

	it('skips hidden columns and shifts the next column left', () => {
		const wb = book({
			cells: { A1: 'a', B1: 'b', C1: 'c' },
			styles: { B1: { fill: solid('FFFF0000') } },
			build: (tx, w) => setColumnsHidden(tx, w, 0, 1, 1, true),
		});
		const l = layout(wb);
		expect(l.texts.map((t) => [t.key, t.cell.x])).toEqual([
			['t:0:0', 0],
			['t:0:2', 64],
		]);
		expect(l.fills).toEqual([]);
	});

	it('lets text overflow across an empty hidden column', () => {
		const wb = book({
			cells: { A1: 'abcdefghijklmnop' },
			build: (tx, w) => setColumnsHidden(tx, w, 0, 1, 1, true),
		});
		// 164 px needed: B adds nothing, C and D add 128 px.
		expect(layout(wb).texts[0].clip).toEqual({ x: 0, y: 0, width: 192, height: 20 });
	});
});

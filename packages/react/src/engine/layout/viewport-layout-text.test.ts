import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	mergeRange,
	parseRange,
	putCell,
	setCellInput,
	setColumnWidth,
	StyleResolver,
} from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import type { DescribeOptions } from '../render/cell-content';
import { setMeasureContextForTests } from '../render/text-measure';
import { MergeIndex } from './style-at';
import type { TextBox, ViewportLayout } from './types';
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

function book(
	cells: Record<string, string>,
	styles: Record<string, StylePatch> = {},
	build?: (tx: EditTransaction, wb: Workbook) => void,
): Workbook {
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

type Options = Partial<Pick<LayoutParams, 'rows' | 'cols' | 'overflowLookaround'>> & {
	zoom?: number;
	describe?: Partial<DescribeOptions>;
};

/** Rows 1-10 and columns A-H at 100%: columns are 64 px, so 6 glyphs (60 px) fit. */
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

function text(l: ViewportLayout, ref: string): TextBox | undefined {
	const { top, left } = parseRange(ref)!;
	return l.texts.find((t) => t.row === top && t.col === left);
}

const LONG = 'abcdefghijklmnopqrstuvwxyz';

describe('layoutViewport: alignment', () => {
	it('aligns text left, numbers right, booleans and errors centre', () => {
		const l = layout(book({ A1: 'abc', B1: '42', C1: 'TRUE', D1: '#N/A' }));
		expect(l.texts.map((t) => [t.key, t.content.kind, t.content.hAlign])).toEqual([
			['t:0:0', 'text', 'left'],
			['t:0:1', 'number', 'right'],
			['t:0:2', 'boolean', 'center'],
			['t:0:3', 'error', 'center'],
		]);
		expect(text(l, 'A1')).toMatchObject({
			cell: { x: 0, y: 0, width: 64, height: 20 },
			hashes: false,
			shrinkScale: 1,
		});
	});

	it('keeps text that fits inside its cell', () => {
		const l = layout(book({ A1: 'abcdef' }));
		expect(text(l, 'A1')?.clip).toEqual(text(l, 'A1')?.cell);
		expect(l.gridlineMasks).toEqual([]);
	});
});

describe('layoutViewport: text overflow', () => {
	it('spills left-aligned text into empty cells on the right and hides their gridlines', () => {
		const l = layout(book({ A1: 'Hello world!', A3: LONG }));
		// 120 px of text + 4 px padding needs one more column; 264 px needs four more.
		expect(text(l, 'A1')?.clip).toEqual({ x: 0, y: 0, width: 128, height: 20 });
		expect(text(l, 'A3')?.clip).toEqual({ x: 0, y: 40, width: 320, height: 20 });
		expect(text(l, 'A1')?.cell).toEqual({ x: 0, y: 0, width: 64, height: 20 });
		expect(l.gridlineMasks).toEqual([
			{ x: 0, y: 0, width: 127, height: 19 },
			{ x: 0, y: 40, width: 319, height: 19 },
		]);
	});

	it('stops at the first cell with a value', () => {
		const l = layout(book({ A1: LONG, C1: 'x' }));
		expect(text(l, 'A1')?.clip.width).toBe(128);
		const blocked = layout(book({ A1: LONG, B1: '0' }));
		expect(text(blocked, 'A1')?.clip.width).toBe(64);
		expect(blocked.gridlineMasks).toEqual([]);
	});

	it('is blocked by formulas, even ones showing nothing, and by merged cells', () => {
		const formula = book({ A1: LONG });
		putCell(formula.sheets[0], 0, 1, { value: '', formula: '""' });
		expect(text(layout(formula), 'A1')?.clip.width).toBe(64);
		const merged = book({ A1: LONG }, {}, (tx, wb) => mergeRange(tx, wb, 0, parseRange('C1:D1')!));
		expect(text(layout(merged), 'A1')?.clip.width).toBe(128);
	});

	it('spills over formatted but empty cells', () => {
		const fill = { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFFF00' } } as const;
		const l = layout(book({ A1: 'Hello world!' }, { B1: { fill } }));
		expect(text(l, 'A1')?.clip.width).toBe(128);
		expect(l.fills.map((f) => f.key)).toEqual(['f:0:1']);
	});

	it('spills right-aligned text to the left and centred text to both sides', () => {
		const right = { alignment: { horizontal: 'right' as const } };
		const center = { alignment: { horizontal: 'center' as const } };
		const l = layout(
			book(
				{ D1: 'Hello world!', D2: 'Hello world!', A3: 'Hello world!' },
				{ D1: right, D2: center, A3: right },
			),
		);
		expect(text(l, 'D1')?.clip).toEqual({ x: 128, y: 0, width: 128, height: 20 });
		// 60 px too wide: 30 px on each side, one column each way.
		expect(text(l, 'D2')?.clip).toEqual({ x: 128, y: 20, width: 192, height: 20 });
		// Nothing left of column A.
		expect(text(l, 'A3')?.clip).toEqual(text(l, 'A3')?.cell);
	});

	it('draws text spilling into view from columns left of the viewport', () => {
		const wb = book({ A1: LONG, A2: 'short' });
		const l = layout(wb, { cols: { start: 2, end: 7 } });
		expect(l.texts.map((t) => t.key)).toEqual(['t:0:0']);
		expect(text(l, 'A1')).toMatchObject({ cell: { x: 0, width: 64 }, clip: { x: 0, width: 320 } });
		expect(l.fills).toEqual([]);
		// Columns beyond the lookaround are not scanned.
		expect(layout(wb, { cols: { start: 2, end: 7 }, overflowLookaround: 1 }).texts).toEqual([]);
	});

	it('keeps wrapped, rotated, vertical and shrunk text inside its cell', () => {
		const l = layout(
			book(
				{ A1: LONG, A2: LONG, A3: LONG, A4: LONG },
				{
					A1: { alignment: { wrapText: true } },
					A2: { alignment: { textRotation: 45 } },
					A3: { alignment: { textRotation: 255 } },
					A4: { alignment: { shrinkToFit: true } },
				},
			),
		);
		for (const ref of ['A1', 'A2', 'A3', 'A4']) {
			expect(text(l, ref)?.clip, ref).toEqual(text(l, ref)?.cell);
		}
		expect(text(l, 'A2')?.content.rotation).toBe(-45);
		expect(text(l, 'A3')?.content.vertical).toBe(true);
		expect(l.gridlineMasks).toEqual([]);
	});

	it('scales the overflow with zoom', () => {
		const l = layout(book({ A1: 'Hello world!' }), { zoom: 2 });
		expect(text(l, 'A1')).toMatchObject({
			cell: { x: 0, y: 0, width: 128, height: 40 },
			clip: { x: 0, y: 0, width: 256, height: 40 },
		});
	});
});

describe('layoutViewport: numbers that do not fit', () => {
	it('shows hashes, as many as fit, instead of spilling', () => {
		const l = layout(
			book({ A1: '1234567.891', B1: '12.5', C1: '123.45' }, { 'A1:C1': { numFmt: '0.00' } }),
		);
		expect(text(l, 'A1')).toMatchObject({ hashes: true, content: { text: '######' } });
		expect(text(l, 'A1')?.clip).toEqual(text(l, 'A1')?.cell);
		expect(text(l, 'B1')).toMatchObject({ hashes: false, content: { text: '12.50' } });
		// Exactly 60 px fits in the 60 px between the paddings.
		expect(text(l, 'C1')).toMatchObject({ hashes: false, content: { text: '123.45' } });
	});

	it('shows hashes for dates and for values the format cannot show', () => {
		const l = layout(
			book(
				{ A1: '1/15/2024', B1: '-1' },
				{ A1: { numFmt: 'mmmm d, yyyy' }, B1: { numFmt: 'm/d/yyyy' } },
				(tx, wb) => setColumnWidth(tx, wb, 0, 1, 1, 20),
			),
		);
		expect(text(l, 'A1')).toMatchObject({ hashes: true, content: { text: '######' } });
		// 140 px wide column: 136 px of room, 13 hashes.
		expect(text(l, 'B1')).toMatchObject({ hashes: true, content: { text: '#'.repeat(13) } });
	});

	it('rounds General numbers to the width before giving up', () => {
		const l = layout(book({ A1: '12345.678', B1: '1234567.891', C1: '-12345.6' }));
		expect(text(l, 'A1')).toMatchObject({ hashes: false, content: { text: '12346' } });
		expect(text(l, 'B1')).toMatchObject({ hashes: false, content: { text: '1E+06' } });
		expect(text(l, 'C1')).toMatchObject({ hashes: false, content: { text: '-12346' } });
		const narrow = layout(
			book({ A1: '12345.678' }, {}, (tx, wb) => setColumnWidth(tx, wb, 0, 0, 0, 1)),
		);
		expect(text(narrow, 'A1')).toMatchObject({ hashes: true, content: { text: '#' } });
	});

	it('leaves room for the indent', () => {
		const l = layout(book({ A1: '123.45' }, { A1: { numFmt: '0.00', alignment: { indent: 1 } } }));
		expect(text(l, 'A1')).toMatchObject({ hashes: true, content: { text: '#####', indent: 9 } });
	});

	// Bug: shrink to fit also applies to numbers in Excel, but layoutViewport checks numbers for
	// hashes first, so 123456.78 in "0.00" with Shrink to fit shows '######' at full size.
	it('shrinks numbers that do not fit when shrink to fit is on', () => {
		const l = layout(
			book({ A1: '123456.78' }, { A1: { numFmt: '0.00', alignment: { shrinkToFit: true } } }),
		);
		expect(text(l, 'A1')).toMatchObject({ hashes: false, content: { text: '123456.78' } });
		expect(text(l, 'A1')?.shrinkScale).toBeCloseTo(60 / 90, 6);
	});

	// Bug: numbers with a repeat-fill (accounting) format are never checked, so 1234567.5 in
	// _($* #,##0.00_) keeps ' $1,234,567.50 ' (150 px) in a 64 px column instead of '#####'.
	it('shows hashes for accounting numbers that do not fit', () => {
		const accounting = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
		const l = layout(book({ A1: '1234567.5' }, { A1: { numFmt: accounting } }));
		expect(text(l, 'A1')?.hashes).toBe(true);
	});
});

describe('layoutViewport: shrink to fit and indent', () => {
	it('scales text down to the room inside the paddings', () => {
		// Typing a line break turns wrapping on, which would disable shrink to fit.
		const shrink = { alignment: { shrinkToFit: true, wrapText: false } };
		const l = layout(
			book(
				{ A1: 'ABCDEFGHIJ', A2: 'abc', A3: 'ab\nABCDEFGHIJKL', A4: 'x'.repeat(1000) },
				{ 'A1:A4': shrink },
			),
		);
		expect(text(l, 'A1')?.shrinkScale).toBeCloseTo(0.6, 9);
		expect(text(l, 'A2')?.shrinkScale).toBe(1);
		// The widest line decides.
		expect(text(l, 'A3')?.shrinkScale).toBeCloseTo(0.5, 9);
		expect(text(l, 'A4')?.shrinkScale).toBe(0.1);
	});

	it('counts the indent in the width text needs', () => {
		const l = layout(book({ A1: 'ABCDE', A2: 'ABCDE' }, { A1: { alignment: { indent: 2 } } }));
		// 50 px + 4 px padding fits in 64 px, plus 18 px of indent it does not.
		expect(text(l, 'A1')).toMatchObject({ content: { indent: 18 }, clip: { width: 128 } });
		expect(text(l, 'A2')?.clip.width).toBe(64);
	});
});

describe('layoutViewport: formula view and zeros', () => {
	it('shows formula text left-aligned, spilling like text', () => {
		const wb = book({ A1: '=SUM(B1:B100)', A2: '=1+2' });
		const l = layout(wb, { describe: { showFormulas: true } });
		expect(text(l, 'A1')).toMatchObject({
			content: { text: '=SUM(B1:B100)', hAlign: 'left' },
			clip: { width: 192 },
		});
		expect(text(l, 'A2')?.content.text).toBe('=1+2');
		// Not calculated yet: nothing to draw in normal view.
		expect(layout(wb).texts).toEqual([]);
	});

	it('hides zero values when asked', () => {
		const wb = book({ A1: '0', B1: '5', C1: '0.0' });
		expect(layout(wb, { describe: { hideZeros: true } }).texts.map((t) => t.key)).toEqual([
			't:0:1',
		]);
		expect(layout(wb).texts).toHaveLength(3);
	});
});

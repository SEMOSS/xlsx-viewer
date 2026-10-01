import { describe, expect, it } from 'vitest';
import type { DifferentialFormat, FillSpec, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	addConditionalFormat,
	applyStylePatch,
	CalcEngine,
	ConditionalFormatEngine,
	createCellIsRule,
	createDataBarRule,
	createTable,
	createWorkbook,
	EditTransaction,
	HIGHLIGHT_PRESETS,
	mergeRange,
	parseRange,
	resolveColorCss,
	setCellInput,
	setColumnsHidden,
	StyleResolver,
} from 'xlsx-viewer-core';

import { SheetGeometry } from '../geometry/sheet-geometry';
import { createCellOverlays } from './overlays';
import { MergeIndex } from './style-at';
import type { CfOverlay, TextBox, ViewportLayout } from './types';
import { layoutViewport, type LayoutParams } from './viewport-layout';

const solid = (rgb: string): FillSpec => ({
	type: 'pattern',
	patternType: 'solid',
	fgColor: { rgb },
});

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

/** A per-cell lookup from values keyed by A1 reference. */
function byRef<T>(values: Record<string, T>): (row: number, col: number) => T | undefined {
	const map = new Map<string, T>();
	for (const [ref, value] of Object.entries(values)) {
		const { top, left } = parseRange(ref)!;
		map.set(`${top}:${left}`, value);
	}
	return (row, col) => map.get(`${row}:${col}`);
}

function layout(
	wb: Workbook,
	extra: Partial<Pick<LayoutParams, 'cf' | 'table'>> = {},
): ViewportLayout {
	const sheet = wb.sheets[0];
	return layoutViewport({
		sheet,
		geometry: new SheetGeometry(sheet, { zoom: 1 }),
		rows: { start: 0, end: 9 },
		cols: { start: 0, end: 7 },
		styles: new StyleResolver(wb.styles, wb.theme),
		merges: new MergeIndex(sheet.merges),
		describe: {
			zoom: 1,
			colorCtx: { theme: wb.theme },
			date1904: false,
			defaultTextColor: '#000000',
		},
		...extra,
	});
}

const fillOf = (l: ViewportLayout, key: string) =>
	l.fills.find((f) => f.key === key)?.css.backgroundColor;
const textOf = (l: ViewportLayout, key: string): TextBox | undefined =>
	l.texts.find((t) => t.key === key);

const BAR = {
	start: 0,
	end: 0.5,
	negative: false,
	color: '#638EC6',
	gradient: true,
	direction: 'ltr',
};

describe('layoutViewport: conditional formatting', () => {
	const wb = book({ A1: '5', B1: '1' }, { A1: { fill: solid('FFFFFF00'), font: { bold: true } } });

	it('draws the rule format over the cell format', () => {
		const cf = byRef<CfOverlay>({ A1: { dxf: HIGHLIGHT_PRESETS.lightRedFillDarkRedText } });
		const l = layout(wb, { cf });
		expect(fillOf(l, 'f:0:0')).toBe('#FFC7CE');
		expect(textOf(l, 't:0:0')?.content).toMatchObject({
			color: '#9C0006',
			font: { fontWeight: 700 },
		});
		const plain = layout(wb);
		expect(fillOf(plain, 'f:0:0')).toBe('#FFFF00');
		expect(textOf(plain, 't:0:0')?.content.color).toBe('#000000');
	});

	it('changes only what the rule sets', () => {
		const red = layout(wb, { cf: byRef<CfOverlay>({ A1: { dxf: HIGHLIGHT_PRESETS.redText } }) });
		expect(fillOf(red, 'f:0:0')).toBe('#FFFF00');
		expect(textOf(red, 't:0:0')?.content.color).toBe('#9C0006');
		const border = layout(wb, {
			cf: byRef<CfOverlay>({ B1: { dxf: HIGHLIGHT_PRESETS.redBorder } }),
		});
		expect(border.borders).toHaveLength(4);
		expect(border.borders.every((b) => b.color === '#9C0006')).toBe(true);
	});

	it('paints color scales over the cell fill and the rule fill', () => {
		const cf = byRef<CfOverlay>({
			A1: { colorScale: '#63BE7B' },
			B1: { colorScale: '#F8696B', dxf: HIGHLIGHT_PRESETS.lightRedFill },
		});
		const l = layout(wb, { cf });
		expect(l.fills.map((f) => [f.key, f.css])).toEqual([
			['f:0:0', { backgroundColor: '#63BE7B' }],
			['f:0:1', { backgroundColor: '#F8696B' }],
		]);
	});

	it('places data bars in the cell box, also when the value is hidden', () => {
		const shown = layout(wb, { cf: byRef<CfOverlay>({ A1: { dataBar: BAR } }) });
		expect(shown.dataBars).toEqual([{ key: 'f:0:0', x: 0, y: 0, width: 64, height: 20, bar: BAR }]);
		expect(shown.dataBars[0].bar).toBe(BAR);
		expect(textOf(shown, 't:0:0')?.content.text).toBe('5');
		const barOnly = layout(wb, { cf: byRef<CfOverlay>({ A1: { dataBar: BAR, hideValue: true } }) });
		expect(barOnly.dataBars).toHaveLength(1);
		expect(barOnly.texts.map((t) => t.key)).toEqual(['t:0:1']);
	});

	it('spans data bars over merged cells and skips hidden columns', () => {
		const merged = book({ A1: '5' }, {}, (tx, w) => {
			mergeRange(tx, w, 0, parseRange('A1:B2')!);
			setColumnsHidden(tx, w, 0, 3, 3, true);
			setCellInput(tx, w, 0, 0, 3, '7');
		});
		const l = layout(merged, {
			cf: byRef<CfOverlay>({ A1: { dataBar: BAR }, D1: { dataBar: BAR } }),
		});
		expect(
			l.dataBars.map(({ key, x, y, width, height }) => ({ key, x, y, width, height })),
		).toEqual([{ key: 'm:0:0', x: 0, y: 0, width: 128, height: 40 }]);
	});

	it('draws icons with the value, or alone when the value is hidden', () => {
		const icon = { iconSet: '3Arrows', index: 2 };
		const l = layout(wb, {
			cf: byRef<CfOverlay>({ A1: { icon }, B1: { icon, hideValue: true }, C1: { icon } }),
		});
		expect(textOf(l, 't:0:0')).toMatchObject({ icon, content: { text: '5', hAlign: 'right' } });
		expect(textOf(l, 't:0:1')).toMatchObject({ icon, content: { text: '' } });
		// No icon without a cell.
		expect(textOf(l, 't:0:2')).toBeUndefined();
	});

	it('uses the overlay of the top-left cell for merged areas', () => {
		const merged = book({ A1: '5' }, {}, (tx, w) => mergeRange(tx, w, 0, parseRange('A1:B2')!));
		const cf = byRef<CfOverlay>({
			A1: { dxf: HIGHLIGHT_PRESETS.greenFillDarkGreenText },
			B2: { colorScale: '#000000' },
		});
		const l = layout(merged, { cf });
		expect(l.fills.map((f) => [f.key, f.css.backgroundColor])).toEqual([['m:0:0', '#C6EFCE']]);
		expect(textOf(l, 't:0:0')?.content.color).toBe('#006100');
	});
});

const ACCENT = { theme: 4 };
const HEADER: DifferentialFormat = {
	font: { bold: true, color: { theme: 0 } },
	fill: { type: 'pattern', patternType: 'solid', bgColor: ACCENT },
	border: { bottom: { style: 'thin', color: ACCENT } },
};
const STRIPE: DifferentialFormat = {
	fill: { type: 'pattern', patternType: 'solid', bgColor: { theme: 4, tint: 0.7999 } },
};

describe('layoutViewport: table formatting', () => {
	const wb = book(
		{ A1: 'Name', B1: 'Qty', A2: 'a', B2: '1' },
		{ B1: { font: { color: { rgb: 'FFFF0000' } } }, B2: { fill: solid('FFFFFF00') } },
	);
	const table = byRef<DifferentialFormat>({ A1: HEADER, B1: HEADER, A2: STRIPE, B2: STRIPE });

	it('formats table cells beneath their own formatting', () => {
		const l = layout(wb, { table });
		expect(l.fills.map((f) => [f.key, f.css.backgroundColor])).toEqual([
			['f:0:0', '#4472C4'],
			['f:0:1', '#4472C4'],
			['f:1:0', '#DAE3F3'],
			['f:1:1', '#FFFF00'],
		]);
		expect(textOf(l, 't:0:0')?.content).toMatchObject({
			color: '#FFFFFF',
			font: { fontWeight: 700 },
		});
		expect(textOf(l, 't:0:1')?.content).toMatchObject({
			color: '#FF0000',
			font: { fontWeight: 700 },
		});
		expect(l.borders.map((b) => [b.orientation, b.at, b.from, b.to, b.color])).toEqual([
			['h', 20, 0, 64, '#4472C4'],
			['h', 20, 64, 128, '#4472C4'],
		]);
	});

	it('puts conditional formatting over table formatting', () => {
		const cf = byRef<CfOverlay>({ A2: { dxf: HIGHLIGHT_PRESETS.lightRedFill } });
		expect(fillOf(layout(wb, { table, cf }), 'f:1:0')).toBe('#FFC7CE');
	});
});

describe('layoutViewport: with the sheet overlays', () => {
	it('combines a styled table, a highlight rule and data bars', () => {
		const wb = book(
			{ A1: 'Name', B1: 'Qty', A2: 'a', B2: '1', A3: 'b', B3: '5', A4: 'c', B4: '10' },
			{},
			(tx, w) => createTable(tx, w, 0, parseRange('A1:B4')!),
		);
		const sheet = wb.sheets[0];
		const qty = parseRange('B2:B4')!;
		addConditionalFormat(
			sheet,
			qty,
			createCellIsRule('greaterThan', ['4'], 'lightRedFill'),
			wb.styles,
		);
		addConditionalFormat(sheet, qty, createDataBarRule('#638EC6'), wb.styles);
		const calc = new CalcEngine(wb);
		const engine = new ConditionalFormatEngine(wb, (f, s, r, c) =>
			calc.evaluateFormula(f, s, r, c),
		);
		const { cf, table } = createCellOverlays(wb, 0, engine);
		const l = layout(wb, { cf, table });

		expect(l.fills.map((f) => [f.key, f.css.backgroundColor])).toEqual([
			['f:0:0', '#4472C4'],
			['f:0:1', '#4472C4'],
			['f:1:0', '#DAE3F3'],
			['f:1:1', '#DAE3F3'],
			['f:2:1', '#FFC7CE'],
			['f:3:0', '#DAE3F3'],
			['f:3:1', '#FFC7CE'],
		]);
		expect(l.dataBars.map((d) => d.key)).toEqual(['f:1:1', 'f:2:1', 'f:3:1']);
		expect(l.dataBars.map((d) => (d.bar as { end: number }).end)).toEqual([0.1, 0.5, 1]);
		expect(textOf(l, 't:0:1')?.content).toMatchObject({
			color: '#FFFFFF',
			font: { fontWeight: 700 },
		});
		// Medium 2 separates the rows with thin accent lines.
		const line = resolveColorCss({ theme: 4, tint: 0.3999 }, { theme: wb.theme });
		expect(l.borders.map((b) => `${b.orientation}${b.at} ${b.from}-${b.to}`).sort()).toEqual(
			['h20 0-64', 'h20 64-128', 'h40 0-64', 'h40 64-128', 'h60 0-64', 'h60 64-128'].sort(),
		);
		expect(l.borders.every((b) => b.color === line)).toBe(true);
	});
});

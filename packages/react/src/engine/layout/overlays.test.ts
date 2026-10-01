import { describe, expect, it } from 'vitest';
import type { DifferentialFormat, ResolvedCellStyle, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	addConditionalFormat,
	CalcEngine,
	ConditionalFormatEngine,
	createCellIsRule,
	createDataBarRule,
	createIconSetRule,
	createTable,
	createWorkbook,
	EditTransaction,
	parseRange,
	setCellInput,
	setTableStyle,
	setTableStyleOptions,
	StylePool,
	StyleResolver,
} from 'xlsx-viewer-core';

import { applyTableDxf, createCellOverlays } from './overlays';

const WHITE = { theme: 0 };
const ACCENT = { theme: 4 };
const HEADER: DifferentialFormat = {
	font: { bold: true, color: WHITE },
	fill: { type: 'pattern', patternType: 'solid', bgColor: ACCENT, fgColor: ACCENT },
	border: { bottom: { style: 'thin', color: ACCENT } },
	numFmt: { id: 10, code: '0.00%' },
};

function styles(...patches: StylePatch[]): {
	normal: ResolvedCellStyle;
	styled: ResolvedCellStyle[];
} {
	const wb = createWorkbook();
	const pool = new StylePool(wb.styles);
	const ids = patches.map((p) => pool.derive(undefined, p));
	const resolver = new StyleResolver(wb.styles, wb.theme);
	return { normal: resolver.resolve(0), styled: ids.map((id) => resolver.resolve(id)) };
}

describe('applyTableDxf', () => {
	it('applies the table format to cells with the Normal style', () => {
		const { normal } = styles();
		const next = applyTableDxf(normal, HEADER, normal);
		expect(next.font).toMatchObject({ bold: true, color: WHITE, name: 'Calibri', size: 11 });
		expect(next.fill).toEqual({ type: 'pattern', patternType: 'solid', fgColor: ACCENT });
		expect(next.border).toEqual({ bottom: { style: 'thin', color: ACCENT } });
		expect([next.numFmt, next.numFmtId]).toEqual(['0.00%', 10]);
	});

	it('lets explicit cell formatting win, property by property', () => {
		const { normal, styled } = styles(
			{ font: { color: { rgb: 'FFFF0000' }, italic: true } },
			{ font: { bold: false } },
			{ fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFFF00' } } },
			{ numFmt: '0.0' },
		);
		const [red, notBold, yellow, decimals] = styled.map((s) => applyTableDxf(s, HEADER, normal));
		expect(red.font).toMatchObject({ color: { rgb: 'FFFF0000' }, italic: true, bold: true });
		expect(notBold.font).toMatchObject({ bold: false, color: WHITE });
		expect(yellow.fill).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { rgb: 'FFFFFF00' },
		});
		expect(yellow.font.bold).toBe(true);
		expect(decimals.numFmt).toBe('0.0');
	});

	it('only fills border edges the cell does not draw', () => {
		const { normal, styled } = styles({
			border: { bottom: { style: 'double', color: { rgb: 'FF00FF00' } } },
		});
		const table: DifferentialFormat = {
			border: { top: { style: 'thin', color: ACCENT }, bottom: { style: 'thin', color: ACCENT } },
		};
		expect(applyTableDxf(styled[0], table, normal).border).toEqual({
			top: { style: 'thin', color: ACCENT },
			bottom: { style: 'double', color: { rgb: 'FF00FF00' } },
		});
		const none = { ...normal, border: { top: { style: 'none' as const } } };
		expect(applyTableDxf(none, table, normal).border.top).toEqual({ style: 'thin', color: ACCENT });
	});

	it('keeps non-solid table fills and leaves the base style untouched', () => {
		const { normal } = styles();
		const grid: DifferentialFormat = {
			fill: { type: 'pattern', patternType: 'lightGrid', fgColor: ACCENT, bgColor: WHITE },
		};
		const snapshot = JSON.stringify(normal);
		expect(applyTableDxf(normal, grid, normal).fill).toBe(grid.fill);
		expect(JSON.stringify(normal)).toBe(snapshot);
	});
});

/** A 2-column table A1:B4 (header Name / Qty, data in rows 2-4) plus the overlays. */
function tableBook(styleName?: string): Workbook {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	const rows = [
		['Name', 'Qty'],
		['a', '1'],
		['b', '5'],
		['c', '10'],
	];
	rows.forEach((cells, r) => cells.forEach((text, c) => setCellInput(tx, wb, 0, r, c, text)));
	createTable(tx, wb, 0, parseRange('A1:B4')!, { styleName });
	tx.commit();
	return wb;
}

function cfEngine(wb: Workbook): ConditionalFormatEngine {
	const calc = new CalcEngine(wb);
	return new ConditionalFormatEngine(wb, (f, s, row, col) => calc.evaluateFormula(f, s, row, col));
}

describe('createCellOverlays: tables', () => {
	it('has no lookups for a sheet without tables or rules', () => {
		const wb = createWorkbook();
		expect(createCellOverlays(wb, 0, null)).toEqual({ table: undefined, cf: undefined });
		expect(createCellOverlays(wb, 0, cfEngine(wb)).cf).toBeUndefined();
	});

	it('formats the header row and banded rows of a TableStyleMedium2 table', () => {
		const { table } = createCellOverlays(tableBook(), 0, null);
		expect(table?.(0, 0)).toMatchObject({
			font: { bold: true, color: WHITE },
			fill: { bgColor: ACCENT },
		});
		expect(table?.(1, 1)?.fill).toMatchObject({ bgColor: { theme: 4, tint: 0.7999 } });
		expect(table?.(2, 1)?.fill).toBeUndefined();
		expect(table?.(3, 0)?.fill).toMatchObject({ bgColor: { theme: 4, tint: 0.7999 } });
		expect(table?.(4, 0)).toBeUndefined();
		expect(table?.(0, 2)).toBeUndefined();
	});

	it('follows the table style and its options', () => {
		const wb = tableBook('TableStyleLight9');
		expect(createCellOverlays(wb, 0, null).table?.(0, 0)?.fill).toMatchObject({ bgColor: ACCENT });
		const tx = new EditTransaction(wb, 'options');
		setTableStyleOptions(tx, wb, 0, 0, { showRowStripes: false, showFirstColumn: true });
		tx.commit();
		const { table } = createCellOverlays(wb, 0, null);
		expect(table?.(1, 0)?.font).toMatchObject({ bold: true });
		expect(table?.(1, 1)?.font).toBeUndefined();
	});

	it('skips tables whose style is unknown', () => {
		expect(createCellOverlays(tableBook('NoSuchStyle'), 0, null).table).toBeUndefined();
	});

	it('skips tables without a style when the workbook has no default table style', () => {
		const wb = tableBook();
		wb.sheets[0].tables[0].style = undefined;
		wb.styles.defaultTableStyle = undefined;
		expect(createCellOverlays(wb, 0, null).table).toBeUndefined();
	});

	// Bug: the gallery's "None" stores the table style without a name (setTableStyle(..., '')),
	// but createCellOverlays treats a missing name as the workbook default (TableStyleMedium2),
	// so a table set to None keeps its blue header and stripes in the grid.
	it('draws no table formatting for the "None" table style', () => {
		const wb = tableBook();
		const tx = new EditTransaction(wb, 'none');
		setTableStyle(tx, wb, 0, 0, '');
		tx.commit();
		expect(wb.sheets[0].tables[0].style?.name).toBeUndefined();
		expect(createCellOverlays(wb, 0, null).table?.(0, 0)).toBeUndefined();
	});
});

describe('createCellOverlays: conditional formatting', () => {
	it('maps engine results to overlays', () => {
		const wb = tableBook();
		const sheet = wb.sheets[0];
		const qty = parseRange('B2:B4')!;
		addConditionalFormat(
			sheet,
			qty,
			createCellIsRule('greaterThan', ['4'], 'lightRedFillDarkRedText'),
			wb.styles,
		);
		addConditionalFormat(sheet, qty, createDataBarRule('#638EC6'), wb.styles);
		addConditionalFormat(sheet, qty, createIconSetRule('3Arrows', { showValue: false }), wb.styles);
		const { cf } = createCellOverlays(wb, 0, cfEngine(wb));
		const low = cf?.(1, 1);
		expect(low?.dxf).toBeUndefined();
		expect(low?.dataBar).toMatchObject({ start: 0, end: 0.1, color: '#638EC6' });
		expect(low?.icon).toEqual({ iconSet: '3Arrows', index: 0 });
		expect(low?.hideValue).toBe(true);
		const high = cf?.(3, 1);
		expect(high?.dxf?.font?.color).toEqual({ rgb: 'FF9C0006' });
		expect(high?.dataBar).toMatchObject({ end: 1 });
		expect(high?.icon?.index).toBe(2);
		expect(cf?.(1, 0)).toBeUndefined();
	});

	it('passes color scales through as CSS colors', () => {
		const wb = tableBook();
		addConditionalFormat(wb.sheets[0], parseRange('B2:B4')!, {
			type: 'colorScale',
			priority: 1,
			colorScale: {
				cfvos: [{ type: 'min' }, { type: 'max' }],
				colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FF000000' }],
			},
		});
		const { cf } = createCellOverlays(wb, 0, cfEngine(wb));
		expect(cf?.(1, 1)?.colorScale).toBe('#FFFFFF');
		expect(cf?.(3, 1)?.colorScale).toBe('#000000');
		expect(cf?.(2, 1)?.colorScale).toMatch(/^#[0-9A-F]{6}$/);
	});
});

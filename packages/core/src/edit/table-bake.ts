import { cellKey } from '../address/cell-address';
import { StyleResolver, type ResolvedCellStyle } from '../styles/resolve-style';
import { StylePool, type StylePatch } from '../styles/style-pool';
import { tableCellFormat } from '../styles/table-format';
import { getTableStyle } from '../styles/table-styles';
import type { ColorSpec, DifferentialFormat, FontSpec } from '../types/style';
import type { Table } from '../types/table';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

function sameColor(a: ColorSpec | undefined, b: ColorSpec | undefined): boolean {
	return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/**
 * Cell formatting that reproduces a table-style format on a cell. Like the
 * renderer, table formatting only shows where the cell keeps the default:
 * font properties equal to the Normal style, no fill, no border edge, the
 * General number format.
 */
export function tableFormatPatch(
	dxf: DifferentialFormat,
	base: ResolvedCellStyle,
	normal: ResolvedCellStyle,
): StylePatch | null {
	const patch: StylePatch = {};
	if (dxf.font) {
		const font: Partial<FontSpec> = {};
		const b = base.font as unknown as Record<string, unknown>;
		const n = normal.font as unknown as Record<string, unknown>;
		for (const [k, v] of Object.entries(dxf.font) as Array<[keyof FontSpec, unknown]>) {
			if (v === undefined) {
				continue;
			}
			const unchanged =
				k === 'color'
					? sameColor(base.font.color, normal.font.color)
					: b[k] === n[k] || (b[k] === undefined && !n[k]);
			if (unchanged) {
				(font as Record<string, unknown>)[k] = v;
			}
		}
		if (Object.keys(font).length > 0) {
			patch.font = font;
		}
	}
	if (dxf.fill && !base.fill) {
		if (dxf.fill.type === 'pattern' && (dxf.fill.patternType ?? 'solid') === 'solid') {
			// Differential solid fills keep their color in bgColor.
			const color = dxf.fill.bgColor ?? dxf.fill.fgColor;
			if (color) {
				patch.fill = { type: 'pattern', patternType: 'solid', fgColor: color };
			}
		} else {
			patch.fill = dxf.fill;
		}
	}
	if (dxf.border) {
		const border: StylePatch['border'] = {};
		for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
			const current = base.border[edge];
			if (dxf.border[edge] && (!current || current.style === 'none')) {
				border[edge] = dxf.border[edge];
			}
		}
		if (Object.keys(border).length > 0) {
			patch.border = border;
		}
	}
	if (dxf.numFmt && base.numFmtId === 0) {
		patch.numFmt = dxf.numFmt.code;
	}
	return Object.keys(patch).length > 0 ? patch : null;
}

/** Apply a table's style to its cells as direct formatting (Convert to Range). */
export function bakeTableStyle(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	table: Table,
): void {
	const style = getTableStyle(table.style?.name, workbook.styles);
	if (!style) {
		return;
	}
	const sheet = workbook.sheets[sheetIndex];
	const resolver = new StyleResolver(workbook.styles, workbook.theme);
	const pool = new StylePool(workbook.styles);
	const normal = resolver.resolve(0);
	tx.touchRange(sheetIndex, table.ref, false);
	for (let row = table.ref.top; row <= table.ref.bottom; row++) {
		for (let col = table.ref.left; col <= table.ref.right; col++) {
			const dxf = tableCellFormat(table, style, row, col);
			if (!dxf) {
				continue;
			}
			const key = cellKey(row, col);
			const cell = sheet.cells.get(key);
			const patch = tableFormatPatch(dxf, resolver.resolve(cell?.style), normal);
			if (patch) {
				sheet.cells.set(key, {
					...(cell ?? { value: null }),
					style: pool.derive(cell?.style ?? 0, patch),
				});
			}
		}
	}
}

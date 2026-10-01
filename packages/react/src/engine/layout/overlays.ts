import type {
	ColorSpec,
	ConditionalFormatEngine,
	DifferentialFormat,
	FontSpec,
	ResolvedCellStyle,
	Table,
	Workbook,
} from 'xlsx-viewer-core';
import { getTableStyle, tableCellFormat } from 'xlsx-viewer-core';

import type { CfOverlay } from './types';

function sameColor(a: ColorSpec | undefined, b: ColorSpec | undefined): boolean {
	if (!a && !b) {
		return true;
	}
	if (!a || !b) {
		return false;
	}
	return (
		a.rgb === b.rgb &&
		a.theme === b.theme &&
		(a.tint ?? 0) === (b.tint ?? 0) &&
		a.indexed === b.indexed &&
		!!a.auto === !!b.auto
	);
}

/**
 * Overlay a table style's formatting beneath a cell's own formatting: a
 * property from the table style only applies where the cell still has the
 * Normal style's value (explicit cell formatting wins, as in Excel).
 */
export function applyTableDxf(
	base: ResolvedCellStyle,
	dxf: DifferentialFormat,
	normal: ResolvedCellStyle,
): ResolvedCellStyle {
	const next: ResolvedCellStyle = { ...base };
	if (dxf.font) {
		const font = { ...base.font };
		const f = font as unknown as Record<string, unknown>;
		const n = normal.font as unknown as Record<string, unknown>;
		for (const [k, v] of Object.entries(dxf.font) as Array<[keyof FontSpec, unknown]>) {
			if (v === undefined) {
				continue;
			}
			const unchanged =
				k === 'color'
					? sameColor(base.font.color, normal.font.color)
					: f[k] === n[k] || (f[k] === undefined && !n[k]);
			if (unchanged) {
				f[k] = v;
			}
		}
		next.font = font;
	}
	if (dxf.fill && !base.fill) {
		if (dxf.fill.type === 'pattern' && (dxf.fill.patternType ?? 'solid') === 'solid') {
			const color = dxf.fill.bgColor ?? dxf.fill.fgColor;
			next.fill = color ? { type: 'pattern', patternType: 'solid', fgColor: color } : undefined;
		} else {
			next.fill = dxf.fill;
		}
	}
	if (dxf.border) {
		const border = { ...base.border };
		for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
			const current = border[edge];
			if (dxf.border[edge] && (!current || current.style === 'none')) {
				border[edge] = dxf.border[edge];
			}
		}
		next.border = border;
	}
	if (dxf.numFmt && base.numFmtId === 0) {
		next.numFmt = dxf.numFmt.code;
		next.numFmtId = dxf.numFmt.id;
	}
	return next;
}

export interface CellOverlays {
	/** Table-style formatting of a cell, if it lies in a styled table. */
	table: ((row: number, col: number) => DifferentialFormat | undefined) | undefined;
	/** Conditional formatting of a cell. */
	cf: ((row: number, col: number) => CfOverlay | undefined) | undefined;
}

/** Per-sheet overlay lookups for the layout (tables and conditional formatting). */
export function createCellOverlays(
	workbook: Workbook,
	sheetIndex: number,
	cfEngine: ConditionalFormatEngine | null,
): CellOverlays {
	const sheet = workbook.sheets[sheetIndex];
	const tables: Array<{ table: Table; style: NonNullable<ReturnType<typeof getTableStyle>> }> = [];
	for (const table of sheet.tables) {
		const style = getTableStyle(table.style?.name, workbook.styles);
		if (style) {
			tables.push({ table, style });
		}
	}
	const table = tables.length
		? (row: number, col: number) => {
				for (const t of tables) {
					const r = t.table.ref;
					if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
						return tableCellFormat(t.table, t.style, row, col);
					}
				}
				return undefined;
			}
		: undefined;
	const cf =
		cfEngine && cfEngine.hasRules(sheetIndex)
			? (row: number, col: number): CfOverlay | undefined => {
					const r = cfEngine.getCellResult(sheetIndex, row, col);
					if (!r) {
						return undefined;
					}
					return {
						dxf: r.style,
						colorScale: r.colorScale,
						dataBar: r.dataBar,
						icon: r.icon,
						hideValue: r.hideValue,
					};
				}
			: undefined;
	return { table, cf };
}

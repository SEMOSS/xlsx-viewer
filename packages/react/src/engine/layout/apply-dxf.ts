import type { DifferentialFormat, ResolvedCellStyle } from 'xlsx-viewer-core';

/**
 * Overlay a differential format (conditional formatting / table styles) on a
 * resolved cell style. Only the properties the dxf specifies change.
 */
export function applyDxf(base: ResolvedCellStyle, dxf: DifferentialFormat): ResolvedCellStyle {
	const next: ResolvedCellStyle = { ...base };
	if (dxf.font) {
		next.font = { ...base.font };
		for (const [k, v] of Object.entries(dxf.font)) {
			if (v !== undefined) {
				(next.font as unknown as Record<string, unknown>)[k] = v;
			}
		}
	}
	if (dxf.fill) {
		// In a dxf, a solid pattern's color is stored as bgColor; normalise to fgColor.
		if (dxf.fill.type === 'pattern') {
			const pattern = dxf.fill.patternType ?? 'solid';
			if (pattern === 'solid' || pattern === 'none') {
				const color = dxf.fill.bgColor ?? dxf.fill.fgColor;
				next.fill = color ? { type: 'pattern', patternType: 'solid', fgColor: color } : base.fill;
			} else {
				next.fill = dxf.fill;
			}
		} else {
			next.fill = dxf.fill;
		}
	}
	if (dxf.border) {
		next.border = { ...base.border };
		for (const [k, v] of Object.entries(dxf.border)) {
			if (v !== undefined) {
				(next.border as unknown as Record<string, unknown>)[k] = v;
			}
		}
	}
	if (dxf.numFmt) {
		next.numFmt = dxf.numFmt.code;
		next.numFmtId = dxf.numFmt.id;
	}
	if (dxf.alignment) {
		next.alignment = { ...base.alignment, ...dxf.alignment };
	}
	return next;
}

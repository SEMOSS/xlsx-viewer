import type { BorderEdge, ColorSpec, DifferentialFormat } from '../types/style';
import type { StyleSheet } from '../types/style';

/** Parts of a table that a table style formats (OOXML `ST_TableStyleType`, table subset). */
export type TableStyleElementType =
	| 'wholeTable'
	| 'headerRow'
	| 'totalRow'
	| 'firstColumn'
	| 'lastColumn'
	| 'firstRowStripe'
	| 'secondRowStripe'
	| 'firstColumnStripe'
	| 'secondColumnStripe'
	| 'firstHeaderCell'
	| 'lastHeaderCell'
	| 'firstTotalCell'
	| 'lastTotalCell';

export interface TableStyleDefinition {
	name: string;
	elements: Partial<Record<TableStyleElementType, DifferentialFormat>>;
	/** Stripe sizes (rows / columns per band), default 1. */
	rowStripeSize?: number;
	colStripeSize?: number;
}

/** Order in which elements are layered (later wins), per ECMA-376 18.8.40. */
export const TABLE_ELEMENT_ORDER: readonly TableStyleElementType[] = [
	'wholeTable',
	'firstColumnStripe',
	'secondColumnStripe',
	'firstRowStripe',
	'secondRowStripe',
	'lastColumn',
	'firstColumn',
	'headerRow',
	'totalRow',
	'firstHeaderCell',
	'lastHeaderCell',
	'firstTotalCell',
	'lastTotalCell',
];

/** Accent per column of each style family: Text 1 then Accent 1-6 (SpreadsheetML theme indices). */
const FAMILY_THEMES = [1, 4, 5, 6, 7, 8, 9];

function themeColor(theme: number, tint?: number): ColorSpec {
	return tint ? { theme, tint } : { theme };
}

function solid(color: ColorSpec): DifferentialFormat['fill'] {
	// Differential fills keep the color in bgColor for solid patterns.
	return { type: 'pattern', patternType: 'solid', bgColor: color, fgColor: color };
}

function edge(color: ColorSpec, style: BorderEdge['style'] = 'thin'): BorderEdge {
	return { style, color };
}

const WHITE: ColorSpec = { theme: 0 };
const BOLD = { bold: true };

function light(index: number): TableStyleDefinition['elements'] {
	// Light 1-7: accent rules above and below, tinted bands; 8-14: accent header; 15-21: full grid.
	const family = Math.floor((index - 1) / 7);
	const t = FAMILY_THEMES[(index - 1) % 7];
	const accent = themeColor(t);
	const band = themeColor(t, t === 1 ? 0.8499 : 0.7999);
	if (family === 0) {
		return {
			wholeTable: {
				border: { top: edge(accent), bottom: edge(accent) },
				font: t === 1 ? undefined : { color: themeColor(t, -0.2499) },
			},
			headerRow: { font: BOLD, border: { bottom: edge(accent) } },
			totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
			firstColumn: { font: BOLD },
			lastColumn: { font: BOLD },
			firstRowStripe: { fill: solid(band) },
			firstColumnStripe: { fill: solid(band) },
		};
	}
	if (family === 1) {
		return {
			wholeTable: {
				border: {
					top: edge(accent),
					bottom: edge(accent),
					left: edge(accent),
					right: edge(accent),
				},
			},
			headerRow: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
			firstColumn: { font: BOLD },
			lastColumn: { font: BOLD },
			firstRowStripe: { border: { top: edge(accent), bottom: edge(accent) } },
			firstColumnStripe: { border: { left: edge(accent), right: edge(accent) } },
		};
	}
	return {
		wholeTable: {
			border: {
				top: edge(accent),
				bottom: edge(accent),
				left: edge(accent),
				right: edge(accent),
				horizontal: edge(accent),
				vertical: edge(accent),
			},
		},
		headerRow: { font: BOLD, border: { bottom: edge(accent) } },
		totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
		firstColumn: { font: BOLD },
		lastColumn: { font: BOLD },
		firstRowStripe: { fill: solid(band) },
		firstColumnStripe: { fill: solid(band) },
	};
}

function medium(index: number): TableStyleDefinition['elements'] {
	const family = Math.floor((index - 1) / 7);
	const t = FAMILY_THEMES[(index - 1) % 7];
	const accent = themeColor(t);
	const band = themeColor(t, t === 1 ? 0.8499 : 0.7999);
	const band2 = themeColor(t, t === 1 ? 0.7499 : 0.5999);
	if (family === 0) {
		// Medium 1-7 (Medium 2 is Excel's default): accent header, tinted bands, no grid.
		return {
			wholeTable: { border: { horizontal: edge(themeColor(t, t === 1 ? 0.7499 : 0.3999)) } },
			headerRow: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
			firstColumn: { font: BOLD },
			lastColumn: { font: BOLD },
			firstRowStripe: { fill: solid(band) },
			firstColumnStripe: { fill: solid(band) },
		};
	}
	if (family === 1) {
		// Medium 8-14: banded tints separated by white grid lines.
		const whiteEdge = edge(WHITE);
		return {
			wholeTable: {
				fill: solid(band),
				font: { color: themeColor(1) },
				border: { horizontal: whiteEdge, vertical: whiteEdge },
			},
			headerRow: {
				font: { bold: true, color: WHITE },
				fill: solid(accent),
				border: { bottom: edge(WHITE, 'thick') },
			},
			totalRow: {
				font: { bold: true, color: WHITE },
				fill: solid(accent),
				border: { top: edge(WHITE, 'thick') },
			},
			firstColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			lastColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			firstRowStripe: { fill: solid(band2) },
			firstColumnStripe: { fill: solid(band2) },
		};
	}
	if (family === 2) {
		// Medium 15-21: dark header with accent rules.
		return {
			wholeTable: {
				border: {
					top: edge(accent, 'medium'),
					bottom: edge(accent, 'medium'),
					horizontal: edge(themeColor(t, 0.3999)),
				},
			},
			headerRow: {
				font: { bold: true, color: WHITE },
				fill: solid(t === 1 ? themeColor(1) : accent),
				border: { bottom: edge(accent, 'medium') },
			},
			totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
			firstColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			lastColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
			firstRowStripe: { fill: solid(themeColor(0, -0.1499)) },
			firstColumnStripe: { fill: solid(themeColor(0, -0.1499)) },
		};
	}
	// Medium 22-28: tinted table with accent grid.
	const gridEdge = edge(themeColor(t, t === 1 ? 0.4999 : 0.3999));
	return {
		wholeTable: {
			fill: solid(band),
			border: {
				top: gridEdge,
				bottom: gridEdge,
				left: gridEdge,
				right: gridEdge,
				horizontal: gridEdge,
				vertical: gridEdge,
			},
		},
		headerRow: { font: BOLD },
		totalRow: { font: BOLD, border: { top: edge(accent, 'double') } },
		firstColumn: { font: BOLD },
		lastColumn: { font: BOLD },
		firstRowStripe: { fill: solid(band2) },
		firstColumnStripe: { fill: solid(band2) },
	};
}

function dark(index: number): TableStyleDefinition['elements'] {
	if (index <= 7) {
		const t = FAMILY_THEMES[(index - 1) % 7];
		const base = themeColor(t, t === 1 ? 0.4999 : -0.2499);
		const stripe = themeColor(t, t === 1 ? 0.3499 : -0.4999);
		return {
			wholeTable: { fill: solid(base), font: { color: WHITE } },
			headerRow: {
				font: { bold: true, color: WHITE },
				fill: solid(themeColor(1)),
				border: { bottom: edge(WHITE, 'medium') },
			},
			totalRow: {
				font: { bold: true, color: WHITE },
				fill: solid(themeColor(1, 0.1499)),
				border: { top: edge(WHITE, 'double') },
			},
			firstColumn: { font: { bold: true, color: WHITE }, fill: solid(stripe) },
			lastColumn: { font: { bold: true, color: WHITE }, fill: solid(stripe) },
			firstRowStripe: { fill: solid(stripe) },
			firstColumnStripe: { fill: solid(stripe) },
		};
	}
	// Dark 8-11: pairs of accents on a light table with a black header.
	const pair = [
		[1, 1],
		[4, 5],
		[6, 7],
		[8, 9],
	][index - 8];
	const accent = themeColor(pair[0]);
	return {
		wholeTable: { fill: solid(themeColor(pair[1], 0.7999)) },
		headerRow: { font: { bold: true, color: WHITE }, fill: solid(themeColor(1)) },
		totalRow: { font: BOLD, border: { top: edge(themeColor(1), 'double') } },
		firstColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
		lastColumn: { font: { bold: true, color: WHITE }, fill: solid(accent) },
		firstRowStripe: { fill: solid(themeColor(pair[1], 0.5999)) },
		firstColumnStripe: { fill: solid(themeColor(pair[1], 0.5999)) },
	};
}

const cache = new Map<string, TableStyleDefinition | null>();

/** Definition of a built-in table style (`TableStyleMedium2`, `TableStyleLight9`, ...). */
export function getBuiltinTableStyle(name: string): TableStyleDefinition | undefined {
	const hit = cache.get(name);
	if (hit !== undefined) {
		return hit ?? undefined;
	}
	const m = /^TableStyle(Light|Medium|Dark)(\d+)$/i.exec(name);
	let def: TableStyleDefinition | null = null;
	if (m) {
		const kind = m[1].toLowerCase();
		const n = Number(m[2]);
		const max = kind === 'light' ? 21 : kind === 'medium' ? 28 : 11;
		if (n >= 1 && n <= max) {
			const elements = kind === 'light' ? light(n) : kind === 'medium' ? medium(n) : dark(n);
			def = { name, elements };
		}
	}
	cache.set(name, def);
	return def ?? undefined;
}

/** Built-in table style names in gallery order. */
export const BUILTIN_TABLE_STYLES: readonly string[] = [
	...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
	...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
	...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];

/**
 * Resolve a table style by name: custom styles from the workbook's styles
 * part (elements point at dxfs) win over built-in names.
 */
export function getTableStyle(
	name: string | undefined,
	styles?: StyleSheet,
): TableStyleDefinition | undefined {
	// A table without a style name uses "None" (the workbook default only applies to new tables).
	if (!name) {
		return undefined;
	}
	const styleName = name;
	const custom = styles?.customTableStyles?.find((s) => s.name === styleName);
	if (custom && styles) {
		const elements: TableStyleDefinition['elements'] = {};
		let rowStripeSize: number | undefined;
		let colStripeSize: number | undefined;
		for (const el of custom.elements) {
			const dxf = styles.dxfs[el.dxfId];
			if (dxf) {
				elements[el.type as TableStyleElementType] = dxf;
			}
			if (el.type === 'firstRowStripe' && el.size) {
				rowStripeSize = el.size;
			}
			if (el.type === 'firstColumnStripe' && el.size) {
				colStripeSize = el.size;
			}
		}
		return { name: styleName, elements, rowStripeSize, colStripeSize };
	}
	return getBuiltinTableStyle(styleName);
}

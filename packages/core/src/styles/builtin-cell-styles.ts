import type {
	BorderSpec,
	CellXf,
	ColorSpec,
	FillSpec,
	FontSpec,
	NamedCellStyle,
	StyleSheet,
} from '../types/style';
import type { Theme } from '../types/theme';
import { StylePool } from './style-pool';

/** Formatting aspects a cell style can include ("Style Includes" in Excel's Style dialog). */
export interface CellStyleIncludes {
	numberFormat: boolean;
	alignment: boolean;
	font: boolean;
	border: boolean;
	fill: boolean;
	protection: boolean;
}

export type CellStyleGroup = 'goodBadNeutral' | 'dataModel' | 'titles' | 'themed' | 'numberFormat';

/** One of Excel's built-in cell styles (Office 2013+ definitions). */
export interface BuiltinCellStyle {
	name: string;
	builtinId: number;
	group: CellStyleGroup;
	/** Properties over the Normal font. */
	font?: Partial<FontSpec>;
	fill?: FillSpec;
	border?: BorderSpec;
	numFmtId?: number;
	includes: CellStyleIncludes;
}

const rgb = (hex: string): ColorSpec => ({ rgb: `FF${hex}` });
const solid = (fgColor: ColorSpec): FillSpec => ({
	type: 'pattern',
	patternType: 'solid',
	fgColor,
	bgColor: { indexed: 64 },
});
const allEdges = (style: 'thin' | 'double', color: ColorSpec): BorderSpec => ({
	left: { style, color },
	right: { style, color },
	top: { style, color },
	bottom: { style, color },
});

const ONLY = (...keys: Array<keyof CellStyleIncludes>): CellStyleIncludes => ({
	numberFormat: keys.includes('numberFormat'),
	alignment: keys.includes('alignment'),
	font: keys.includes('font'),
	border: keys.includes('border'),
	fill: keys.includes('fill'),
	protection: keys.includes('protection'),
});

const ALL = ONLY('numberFormat', 'alignment', 'font', 'border', 'fill', 'protection');
const FONT_FILL = ONLY('font', 'fill');
const FONT_FILL_BORDER = ONLY('font', 'fill', 'border');

/** Tints of the "20% / 40% / 60% - Accent" styles. */
const ACCENT_TINTS = {
	20: 0.7999816888943144,
	40: 0.5999938962981048,
	60: 0.3999755851924192,
} as const;

function accentStyles(): BuiltinCellStyle[] {
	const out: BuiltinCellStyle[] = [];
	for (let n = 1; n <= 6; n++) {
		const theme = 3 + n;
		const base = 29 + (n - 1) * 4;
		const tinted = (pct: 20 | 40 | 60, id: number): BuiltinCellStyle => ({
			name: `${pct}% - Accent${n}`,
			builtinId: id,
			group: 'themed',
			font: { color: { theme: pct === 60 ? 0 : 1 } },
			fill: solid({ theme, tint: ACCENT_TINTS[pct] }),
			includes: FONT_FILL,
		});
		out.push(tinted(20, base + 1), tinted(40, base + 2), tinted(60, base + 3), {
			name: `Accent${n}`,
			builtinId: base,
			group: 'themed',
			font: { color: { theme: 0 } },
			fill: solid({ theme }),
			includes: FONT_FILL,
		});
	}
	return out;
}

/** Excel's built-in cell styles in gallery order. */
export const BUILTIN_CELL_STYLES: readonly BuiltinCellStyle[] = [
	{ name: 'Normal', builtinId: 0, group: 'goodBadNeutral', includes: ALL },
	{
		name: 'Bad',
		builtinId: 27,
		group: 'goodBadNeutral',
		font: { color: rgb('9C0006') },
		fill: solid(rgb('FFC7CE')),
		includes: FONT_FILL,
	},
	{
		name: 'Good',
		builtinId: 26,
		group: 'goodBadNeutral',
		font: { color: rgb('006100') },
		fill: solid(rgb('C6EFCE')),
		includes: FONT_FILL,
	},
	{
		name: 'Neutral',
		builtinId: 28,
		group: 'goodBadNeutral',
		font: { color: rgb('9C5700') },
		fill: solid(rgb('FFEB9C')),
		includes: FONT_FILL,
	},
	{
		name: 'Calculation',
		builtinId: 22,
		group: 'dataModel',
		font: { bold: true, color: rgb('FA7D00') },
		fill: solid(rgb('F2F2F2')),
		border: allEdges('thin', rgb('7F7F7F')),
		includes: FONT_FILL_BORDER,
	},
	{
		name: 'Check Cell',
		builtinId: 23,
		group: 'dataModel',
		font: { bold: true, color: { theme: 0 } },
		fill: solid(rgb('A5A5A5')),
		border: allEdges('double', rgb('3F3F3F')),
		includes: FONT_FILL_BORDER,
	},
	{
		name: 'Explanatory Text',
		builtinId: 53,
		group: 'dataModel',
		font: { italic: true, color: rgb('7F7F7F') },
		includes: ONLY('font'),
	},
	{
		name: 'Input',
		builtinId: 20,
		group: 'dataModel',
		font: { color: rgb('3F3F76') },
		fill: solid(rgb('FFCC99')),
		border: allEdges('thin', rgb('7F7F7F')),
		includes: FONT_FILL_BORDER,
	},
	{
		name: 'Linked Cell',
		builtinId: 24,
		group: 'dataModel',
		font: { color: rgb('FA7D00') },
		border: { bottom: { style: 'double', color: rgb('FF8001') } },
		includes: ONLY('font', 'border'),
	},
	{
		name: 'Note',
		builtinId: 10,
		group: 'dataModel',
		fill: solid(rgb('FFFFCC')),
		border: allEdges('thin', rgb('B2B2B2')),
		includes: ONLY('fill', 'border'),
	},
	{
		name: 'Output',
		builtinId: 21,
		group: 'dataModel',
		font: { bold: true, color: rgb('3F3F3F') },
		fill: solid(rgb('F2F2F2')),
		border: allEdges('thin', rgb('3F3F3F')),
		includes: FONT_FILL_BORDER,
	},
	{
		name: 'Warning Text',
		builtinId: 11,
		group: 'dataModel',
		font: { color: rgb('FF0000') },
		includes: ONLY('font'),
	},
	{
		name: 'Heading 1',
		builtinId: 16,
		group: 'titles',
		font: { bold: true, size: 15, color: { theme: 3 } },
		border: { bottom: { style: 'thick', color: { theme: 4 } } },
		includes: ONLY('font', 'border'),
	},
	{
		name: 'Heading 2',
		builtinId: 17,
		group: 'titles',
		font: { bold: true, size: 13, color: { theme: 3 } },
		border: { bottom: { style: 'thick', color: { theme: 4, tint: 0.499984740745262 } } },
		includes: ONLY('font', 'border'),
	},
	{
		name: 'Heading 3',
		builtinId: 18,
		group: 'titles',
		font: { bold: true, color: { theme: 3 } },
		border: { bottom: { style: 'medium', color: { theme: 4, tint: 0.3999755851924192 } } },
		includes: ONLY('font', 'border'),
	},
	{
		name: 'Heading 4',
		builtinId: 19,
		group: 'titles',
		font: { bold: true, color: { theme: 3 } },
		includes: ONLY('font'),
	},
	{
		name: 'Title',
		builtinId: 15,
		group: 'titles',
		font: { size: 18, color: { theme: 3 }, scheme: 'major' },
		includes: ONLY('font'),
	},
	{
		name: 'Total',
		builtinId: 25,
		group: 'titles',
		font: { bold: true, color: { theme: 1 } },
		border: {
			top: { style: 'thin', color: { theme: 4 } },
			bottom: { style: 'double', color: { theme: 4 } },
		},
		includes: ONLY('font', 'border'),
	},
	...accentStyles(),
	{
		name: 'Comma',
		builtinId: 3,
		group: 'numberFormat',
		numFmtId: 43,
		includes: ONLY('numberFormat'),
	},
	{
		name: 'Comma [0]',
		builtinId: 6,
		group: 'numberFormat',
		numFmtId: 41,
		includes: ONLY('numberFormat'),
	},
	{
		name: 'Currency',
		builtinId: 4,
		group: 'numberFormat',
		numFmtId: 44,
		includes: ONLY('numberFormat'),
	},
	{
		name: 'Currency [0]',
		builtinId: 7,
		group: 'numberFormat',
		numFmtId: 42,
		includes: ONLY('numberFormat'),
	},
	{
		name: 'Percent',
		builtinId: 5,
		group: 'numberFormat',
		numFmtId: 9,
		includes: ONLY('numberFormat'),
	},
];

/** A built-in cell style by name (case-insensitive). */
export function getBuiltinCellStyle(name: string): BuiltinCellStyle | undefined {
	const lower = name.toLowerCase();
	return BUILTIN_CELL_STYLES.find((s) => s.name.toLowerCase() === lower);
}

/** The workbook's named style with this name (or this built-in id), if any. */
export function findCellStyle(styles: StyleSheet, name: string): NamedCellStyle | undefined {
	const lower = name.toLowerCase();
	const builtin = getBuiltinCellStyle(name);
	return (
		styles.cellStyles.find((s) => s.name.toLowerCase() === lower) ??
		(builtin
			? styles.cellStyles.find((s) => s.builtinId === builtin.builtinId && !s.customBuiltin)
			: undefined)
	);
}

/** What a named style's cellStyleXfs entry includes (absent `apply*` flags mean included). */
export function cellStyleIncludes(xf: CellXf | undefined): CellStyleIncludes {
	return {
		numberFormat: xf?.applyNumberFormat !== false,
		alignment: xf?.applyAlignment !== false,
		font: xf?.applyFont !== false,
		border: xf?.applyBorder !== false,
		fill: xf?.applyFill !== false,
		protection: xf?.applyProtection !== false,
	};
}

/** The font a built-in style uses: the Normal font with the style's changes (major-scheme fonts get the theme name). */
export function builtinCellStyleFont(
	styles: StyleSheet,
	def: BuiltinCellStyle,
	theme?: Theme,
): FontSpec {
	const normal = styles.fonts[0] ?? {};
	const font: FontSpec = { ...normal, ...(def.font ?? {}) };
	if (def.font?.scheme === 'major') {
		font.name = theme?.fonts.majorLatin ?? 'Calibri Light';
	}
	return font;
}

/**
 * The cellStyleXfs index of a named style, adding Excel's built-in
 * definition when the workbook lacks it. Only appends (fonts, fills,
 * borders, cellStyleXfs, cellStyles). Returns undefined for unknown names.
 */
export function ensureCellStyle(
	styles: StyleSheet,
	name: string,
	theme?: Theme,
): number | undefined {
	const existing = findCellStyle(styles, name);
	if (existing) {
		return existing.xfId;
	}
	const def = getBuiltinCellStyle(name);
	if (!def) {
		return undefined;
	}
	const pool = new StylePool(styles);
	const normal = styles.cellStyleXfs[0] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 };
	const inc = def.includes;
	styles.cellStyleXfs.push({
		numFmtId: def.numFmtId ?? normal.numFmtId,
		fontId: def.font ? pool.internFont(builtinCellStyleFont(styles, def, theme)) : normal.fontId,
		fillId: def.fill ? pool.internFill(def.fill) : normal.fillId,
		borderId: def.border ? pool.internBorder(def.border) : normal.borderId,
		applyNumberFormat: inc.numberFormat ? undefined : false,
		applyFont: inc.font ? undefined : false,
		applyFill: inc.fill ? undefined : false,
		applyBorder: inc.border ? undefined : false,
		applyAlignment: inc.alignment ? undefined : false,
		applyProtection: inc.protection ? undefined : false,
	});
	const xfId = styles.cellStyleXfs.length - 1;
	styles.cellStyles.push({ name: def.name, xfId, builtinId: def.builtinId });
	return xfId;
}

/**
 * The cellXfs index of `baseXfId` restyled with a named style: the aspects
 * the style includes come from the style, the rest stay as they were.
 */
export function applyNamedStyleToXf(
	styles: StyleSheet,
	baseXfId: number | undefined,
	styleXfId: number,
): number {
	const pool = new StylePool(styles);
	const base: CellXf = styles.cellXfs[baseXfId ?? 0] ??
		styles.cellXfs[0] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 };
	const style = styles.cellStyleXfs[styleXfId];
	const inc = cellStyleIncludes(style);
	const next: CellXf = { ...base, xfId: styleXfId };
	if (style) {
		if (inc.numberFormat) {
			next.numFmtId = style.numFmtId;
			next.applyNumberFormat = undefined;
		}
		if (inc.font) {
			next.fontId = style.fontId;
			next.applyFont = undefined;
		}
		if (inc.fill) {
			next.fillId = style.fillId;
			next.applyFill = undefined;
		}
		if (inc.border) {
			next.borderId = style.borderId;
			next.applyBorder = undefined;
		}
		if (inc.alignment) {
			next.alignment = style.alignment;
			next.applyAlignment = undefined;
		}
		if (inc.protection) {
			next.protection = style.protection;
			next.applyProtection = undefined;
		}
	}
	return pool.internXf(next);
}

import { StylePool } from '../styles/style-pool';
import type { FontSpec, StyleSheet } from '../types/style';

/**
 * Excel's built-in "Hyperlink" cell style (builtinId 8): the Normal font
 * in the theme's hyperlink color (theme color 10), single underlined.
 * Entries are only ever appended, never modified.
 */

/** `builtinId` of the Hyperlink cell style. */
export const HYPERLINK_BUILTIN_STYLE_ID = 8;

/** The cellStyleXfs index of the Hyperlink cell style, when the workbook has one. */
export function findHyperlinkCellStyle(styles: StyleSheet): number | undefined {
	const style =
		styles.cellStyles.find((s) => s.builtinId === HYPERLINK_BUILTIN_STYLE_ID) ??
		styles.cellStyles.find((s) => s.name.toLowerCase() === 'hyperlink');
	return style?.xfId;
}

function hyperlinkFont(styles: StyleSheet): FontSpec {
	return { ...(styles.fonts[0] ?? {}), color: { theme: 10 }, underline: 'single' };
}

/**
 * The cellStyleXfs index of the Hyperlink cell style, creating it when
 * missing: a font (Normal font, theme color 10, single underline), a
 * cellStyleXfs entry using it and a `Hyperlink` cellStyles entry.
 */
export function ensureHyperlinkCellStyle(styles: StyleSheet): number {
	const existing = findHyperlinkCellStyle(styles);
	if (existing !== undefined) {
		return existing;
	}
	const fontId = new StylePool(styles).internFont(hyperlinkFont(styles));
	styles.cellStyleXfs.push({
		numFmtId: 0,
		fontId,
		fillId: 0,
		borderId: 0,
		applyNumberFormat: false,
		applyFill: false,
		applyBorder: false,
		applyAlignment: false,
		applyProtection: false,
	});
	const xfId = styles.cellStyleXfs.length - 1;
	styles.cellStyles.push({ name: 'Hyperlink', xfId, builtinId: HYPERLINK_BUILTIN_STYLE_ID });
	return xfId;
}

/** The cellXfs index for a cell formatted with the Hyperlink style (shared or appended). */
export function hyperlinkCellXf(styles: StyleSheet): number {
	const xfId = ensureHyperlinkCellStyle(styles);
	const pool = new StylePool(styles);
	const fontId = styles.cellStyleXfs[xfId]?.fontId ?? pool.internFont(hyperlinkFont(styles));
	return pool.internXf({ numFmtId: 0, fontId, fillId: 0, borderId: 0, xfId, applyFont: true });
}

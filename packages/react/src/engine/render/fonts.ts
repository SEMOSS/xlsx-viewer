import type { FontSpec } from 'xlsx-viewer-core';

/**
 * Metric-compatible fallbacks for common Office fonts, so text lays out at
 * the same width on machines without the Microsoft fonts installed.
 */
const FONT_FALLBACKS: Record<string, string> = {
	calibri: "Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif",
	'calibri light':
		"'Calibri Light', Calibri, Carlito, 'Segoe UI Light', 'Helvetica Neue', Arial, sans-serif",
	aptos:
		"Aptos, 'Aptos Display', Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif",
	'aptos narrow':
		"'Aptos Narrow', Aptos, 'Arial Narrow', Calibri, Carlito, 'Helvetica Neue', Arial, sans-serif",
	'aptos display': "'Aptos Display', Aptos, Calibri, Carlito, 'Helvetica Neue', Arial, sans-serif",
	arial: "Arial, 'Liberation Sans', Arimo, 'Helvetica Neue', Helvetica, sans-serif",
	'arial narrow': "'Arial Narrow', 'Liberation Sans Narrow', Arial, sans-serif",
	helvetica: "Helvetica, 'Helvetica Neue', Arial, 'Liberation Sans', sans-serif",
	'times new roman': "'Times New Roman', 'Liberation Serif', Tinos, Times, serif",
	cambria: "Cambria, Caladea, 'Times New Roman', serif",
	georgia: 'Georgia, Gelasio, serif',
	garamond: "Garamond, 'EB Garamond', 'Times New Roman', serif",
	'courier new': "'Courier New', 'Liberation Mono', Cousine, Courier, monospace",
	consolas: "Consolas, 'Cascadia Mono', Menlo, 'DejaVu Sans Mono', monospace",
	verdana: "Verdana, 'DejaVu Sans', sans-serif",
	tahoma: "Tahoma, 'DejaVu Sans', Verdana, sans-serif",
	'segoe ui': "'Segoe UI', system-ui, -apple-system, 'Helvetica Neue', Arial, sans-serif",
	'trebuchet ms': "'Trebuchet MS', 'DejaVu Sans', sans-serif",
	'century gothic': "'Century Gothic', 'URW Gothic', 'Avenir Next', sans-serif",
	'franklin gothic medium': "'Franklin Gothic Medium', 'Libre Franklin', Arial, sans-serif",
	'meiryo ui': "'Meiryo UI', Meiryo, 'Hiragino Sans', sans-serif",
	'ms gothic': "'MS Gothic', 'Hiragino Kaku Gothic ProN', monospace",
	'ms pgothic': "'MS PGothic', 'Hiragino Kaku Gothic ProN', sans-serif",
	'yu gothic': "'Yu Gothic', YuGothic, 'Hiragino Sans', sans-serif",
	simsun: "SimSun, 'Songti SC', serif",
	'microsoft yahei': "'Microsoft YaHei', 'PingFang SC', sans-serif",
	'malgun gothic': "'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif",
};

function quoteFamily(name: string): string {
	return /^[A-Za-z-]+$/.test(name) ? name : `'${name.replace(/'/g, "\\'")}'`;
}

/** CSS `font-family` stack for an Excel font name (and OOXML family class). */
export function fontFamilyStack(name: string | undefined, familyClass?: number): string {
	const key = (name ?? 'Calibri').trim().toLowerCase();
	const known = FONT_FALLBACKS[key];
	if (known) {
		return known;
	}
	const generic =
		familyClass === 1
			? 'serif'
			: familyClass === 3
				? 'monospace'
				: familyClass === 4
					? 'cursive'
					: 'sans-serif';
	return `${quoteFamily(name ?? 'Calibri')}, Calibri, Carlito, 'Helvetica Neue', Arial, ${generic}`;
}

/** Point size to CSS pixels at a zoom level. */
export function fontSizePx(sizePt: number | undefined, zoom: number): number {
	return ((sizePt ?? 11) * 96 * zoom) / 72;
}

export interface FontCss {
	fontFamily: string;
	/** Pixels. */
	fontSize: number;
	fontWeight: number;
	fontStyle: 'normal' | 'italic';
	/** CSS `text-decoration-line`. */
	textDecorationLine?: string;
	textDecorationStyle?: 'double';
	/** Superscript / subscript. */
	verticalAlign?: 'super' | 'sub';
}

/** CSS font properties for a font spec. Colors are resolved separately. */
export function fontToCss(font: FontSpec, zoom: number): FontCss {
	const decorations: string[] = [];
	const underline = font.underline && font.underline !== 'none';
	if (underline) {
		decorations.push('underline');
	}
	if (font.strike) {
		decorations.push('line-through');
	}
	const scriptScale = font.vertAlign === 'superscript' || font.vertAlign === 'subscript' ? 0.7 : 1;
	return {
		fontFamily: fontFamilyStack(font.name, font.family),
		fontSize: fontSizePx(font.size, zoom) * scriptScale,
		fontWeight: font.bold ? 700 : 400,
		fontStyle: font.italic ? 'italic' : 'normal',
		textDecorationLine: decorations.length ? decorations.join(' ') : undefined,
		textDecorationStyle:
			font.underline === 'double' || font.underline === 'doubleAccounting' ? 'double' : undefined,
		verticalAlign:
			font.vertAlign === 'superscript'
				? 'super'
				: font.vertAlign === 'subscript'
					? 'sub'
					: undefined,
	};
}

/** Canvas `font` shorthand for measuring text. */
export function canvasFont(css: FontCss): string {
	return `${css.fontStyle} ${css.fontWeight} ${css.fontSize}px ${css.fontFamily}`;
}

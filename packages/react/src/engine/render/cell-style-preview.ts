import type {
	BorderEdge,
	BorderSpec,
	ColorContext,
	FillSpec,
	FontSpec,
	StyleSheet,
	Theme,
} from 'xlsx-viewer-core';
import {
	builtinCellStyleFont,
	findCellStyle,
	getBuiltinCellStyle,
	resolveColorCss,
} from 'xlsx-viewer-core';

import { fillToCss } from './fills';
import { fontFamilyStack } from './fonts';
import { borderStroke } from './borders';

/** CSS for a Cell Styles gallery tile: the style's name drawn in the style. */
export interface CellStylePreview {
	fontFamily: string;
	/** Px, capped so big headings still fit a tile. */
	fontSize: number;
	fontWeight: number;
	fontStyle: 'normal' | 'italic';
	color: string;
	backgroundColor?: string;
	backgroundImage?: string;
	borderTop?: string;
	borderBottom?: string;
	borderLeft?: string;
	borderRight?: string;
}

function edgeCss(edge: BorderEdge | undefined, ctx: ColorContext): string | undefined {
	const stroke = edge ? borderStroke(edge.style) : undefined;
	if (!edge || !stroke) {
		return undefined;
	}
	const color = resolveColorCss(edge.color, ctx, '#000000') ?? '#000000';
	const kind = stroke.double
		? 'double'
		: stroke.dash
			? edge.style === 'dotted' || edge.style === 'hair'
				? 'dotted'
				: 'dashed'
			: 'solid';
	return `${stroke.width}px ${kind} ${color}`;
}

/**
 * Preview of a named cell style: the workbook's own definition when it has
 * the style, else Excel's built-in one. Undefined for unknown names.
 */
export function cellStylePreview(
	styles: StyleSheet,
	theme: Theme,
	name: string,
	maxFontPx = 20,
): CellStylePreview | undefined {
	const ctx: ColorContext = { theme, indexedColors: styles.indexedColors };
	let font: FontSpec;
	let fill: FillSpec | undefined;
	let border: BorderSpec | undefined;
	const named = findCellStyle(styles, name);
	const xf = named ? styles.cellStyleXfs[named.xfId] : undefined;
	if (xf) {
		font = styles.fonts[xf.fontId] ?? styles.fonts[0] ?? {};
		fill = styles.fills[xf.fillId];
		border = styles.borders[xf.borderId];
	} else {
		const def = getBuiltinCellStyle(name);
		if (!def) {
			return undefined;
		}
		font = builtinCellStyleFont(styles, def, theme);
		fill = def.fill;
		border = def.border;
	}
	const scheme =
		font.scheme === 'major'
			? theme?.fonts.majorLatin
			: font.scheme === 'minor'
				? theme?.fonts.minorLatin
				: undefined;
	const css = fillToCss(fill, ctx);
	return {
		fontFamily: fontFamilyStack(scheme ?? font.name, font.family),
		fontSize: Math.min(maxFontPx, Math.round(((font.size ?? 11) * 96) / 72)),
		fontWeight: font.bold ? 700 : 400,
		fontStyle: font.italic ? 'italic' : 'normal',
		color: resolveColorCss(font.color, ctx, '#000000') ?? '#000000',
		backgroundColor: css?.backgroundColor,
		backgroundImage: css?.backgroundImage,
		borderTop: edgeCss(border?.top, ctx),
		borderBottom: edgeCss(border?.bottom, ctx),
		borderLeft: edgeCss(border?.left, ctx),
		borderRight: edgeCss(border?.right, ctx),
	};
}

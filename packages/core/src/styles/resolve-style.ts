import type {
	AlignmentSpec,
	BorderSpec,
	CellXf,
	FillSpec,
	FontSpec,
	ProtectionSpec,
	StyleSheet,
} from '../types/style';
import type { Theme } from '../types/theme';
import { numberFormatCode } from './builtin-number-formats';

/** Everything needed to draw a cell, with indices dereferenced. */
export interface ResolvedCellStyle {
	xfId: number;
	/** Font with name and size filled in from the Normal style when missing. */
	font: FontSpec & { name: string; size: number };
	fill?: FillSpec;
	border: BorderSpec;
	alignment: AlignmentSpec;
	numFmtId: number;
	numFmt: string;
	protection: Required<ProtectionSpec>;
	quotePrefix: boolean;
}

const EMPTY_BORDER: BorderSpec = {};
const EMPTY_ALIGNMENT: AlignmentSpec = {};

/** Resolve the font family to use, mapping theme scheme fonts to the theme's fonts. */
export function effectiveFontName(
	font: FontSpec,
	theme: Theme | undefined,
	fallback: string,
): string {
	if (theme && font.scheme === 'minor' && theme.fonts.minorLatin) {
		return theme.fonts.minorLatin;
	}
	if (theme && font.scheme === 'major' && theme.fonts.majorLatin) {
		return theme.fonts.majorLatin;
	}
	return font.name ?? fallback;
}

/**
 * Caches resolved styles per cellXfs index. Invalidate (or create a new
 * resolver) whenever the stylesheet arrays grow.
 */
export class StyleResolver {
	private readonly cache = new Map<number, ResolvedCellStyle>();
	private readonly defaultFont: FontSpec & { name: string; size: number };

	constructor(
		readonly styles: StyleSheet,
		readonly theme?: Theme,
	) {
		const f0 = styles.fonts[0] ?? {};
		this.defaultFont = {
			...f0,
			name: effectiveFontName(f0, theme, 'Calibri'),
			size: f0.size ?? 11,
		};
	}

	/** The Normal style font (fonts[0]). */
	get normalFont(): FontSpec & { name: string; size: number } {
		return this.defaultFont;
	}

	resolve(xfId: number | undefined): ResolvedCellStyle {
		const id = xfId ?? 0;
		const hit = this.cache.get(id);
		if (hit) {
			return hit;
		}
		const xf: CellXf = this.styles.cellXfs[id] ??
			this.styles.cellXfs[0] ?? {
				numFmtId: 0,
				fontId: 0,
				fillId: 0,
				borderId: 0,
			};
		const f = this.styles.fonts[xf.fontId] ?? {};
		const font = {
			...this.defaultFont,
			...f,
			name: effectiveFontName(f, this.theme, this.defaultFont.name),
			size: f.size ?? this.defaultFont.size,
		};
		const fill = this.styles.fills[xf.fillId];
		const resolved: ResolvedCellStyle = {
			xfId: id,
			font,
			fill: fill && !(fill.type === 'pattern' && fill.patternType === 'none') ? fill : undefined,
			border: this.styles.borders[xf.borderId] ?? EMPTY_BORDER,
			alignment: xf.alignment ?? EMPTY_ALIGNMENT,
			numFmtId: xf.numFmtId,
			numFmt: numberFormatCode(xf.numFmtId, this.styles.numFmts),
			protection: {
				locked: xf.protection?.locked ?? true,
				hidden: xf.protection?.hidden ?? false,
			},
			quotePrefix: xf.quotePrefix ?? false,
		};
		this.cache.set(id, resolved);
		return resolved;
	}

	/** Drop cached entries (call after the stylesheet changes). */
	clear(): void {
		this.cache.clear();
	}
}

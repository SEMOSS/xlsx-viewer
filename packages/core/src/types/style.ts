import type { XmlElement } from '../xml/types';

/**
 * Styles model, intentionally close to SpreadsheetML `styles.xml` so that
 * indices stay stable across load and save. Collections are **append-only**
 * while editing: existing entries are never reordered or mutated, so cells in
 * untouched sheets keep pointing at the right formatting.
 */

/** OOXML `CT_Color`. At most one of `rgb`, `theme`, `indexed`, `auto` is normally set. */
export interface ColorSpec {
	/** ARGB hex, e.g. `FF0070C0`. */
	rgb?: string;
	/** Theme color index (0 = lt1/Background 1, 1 = dk1/Text 1, 2 = lt2, 3 = dk2, 4-9 = accent1-6, 10 = hlink, 11 = folHlink). */
	theme?: number;
	/** Tint in [-1, 1] applied to the base color. */
	tint?: number;
	/** Index into the legacy indexed palette (64 = system foreground, 65 = system background). */
	indexed?: number;
	/** Automatic (system) color. */
	auto?: boolean;
}

export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none';

export interface FontSpec {
	name?: string;
	/** Point size. */
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: UnderlineStyle;
	strike?: boolean;
	color?: ColorSpec;
	vertAlign?: 'superscript' | 'subscript' | 'baseline';
	/** Font family class (1 = Roman, 2 = Swiss, 3 = Modern, 4 = Script, 5 = Decorative). */
	family?: number;
	charset?: number;
	scheme?: 'major' | 'minor' | 'none';
	outline?: boolean;
	shadow?: boolean;
	condense?: boolean;
	extend?: boolean;
}

export type PatternType =
	| 'none'
	| 'solid'
	| 'mediumGray'
	| 'darkGray'
	| 'lightGray'
	| 'darkHorizontal'
	| 'darkVertical'
	| 'darkDown'
	| 'darkUp'
	| 'darkGrid'
	| 'darkTrellis'
	| 'lightHorizontal'
	| 'lightVertical'
	| 'lightDown'
	| 'lightUp'
	| 'lightGrid'
	| 'lightTrellis'
	| 'gray125'
	| 'gray0625';

export interface PatternFill {
	type: 'pattern';
	patternType: PatternType;
	/** Foreground (pattern) color; for `solid` this is the cell background. */
	fgColor?: ColorSpec;
	bgColor?: ColorSpec;
}

export interface GradientStop {
	/** Position in [0, 1]. */
	position: number;
	color: ColorSpec;
}

export interface GradientFill {
	type: 'gradient';
	gradientType: 'linear' | 'path';
	/** Angle in degrees for linear gradients. */
	degree?: number;
	left?: number;
	right?: number;
	top?: number;
	bottom?: number;
	stops: GradientStop[];
}

export type FillSpec = PatternFill | GradientFill;

export type BorderStyle =
	| 'none'
	| 'thin'
	| 'medium'
	| 'dashed'
	| 'dotted'
	| 'thick'
	| 'double'
	| 'hair'
	| 'mediumDashed'
	| 'dashDot'
	| 'mediumDashDot'
	| 'dashDotDot'
	| 'mediumDashDotDot'
	| 'slantDashDot';

export interface BorderEdge {
	style: BorderStyle;
	color?: ColorSpec;
}

export interface BorderSpec {
	left?: BorderEdge;
	right?: BorderEdge;
	top?: BorderEdge;
	bottom?: BorderEdge;
	diagonal?: BorderEdge;
	diagonalUp?: boolean;
	diagonalDown?: boolean;
	/** Inner vertical / horizontal edges (only meaningful in differential formats). */
	vertical?: BorderEdge;
	horizontal?: BorderEdge;
	outline?: boolean;
}

export type HorizontalAlignment =
	| 'general'
	| 'left'
	| 'center'
	| 'right'
	| 'fill'
	| 'justify'
	| 'centerContinuous'
	| 'distributed';

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface AlignmentSpec {
	horizontal?: HorizontalAlignment;
	vertical?: VerticalAlignment;
	wrapText?: boolean;
	shrinkToFit?: boolean;
	/** Indent level (each level is about three spaces of the default font). */
	indent?: number;
	/** 0-90 counter-clockwise, 91-180 clockwise (value - 90), 255 = vertical stacked text. */
	textRotation?: number;
	/** 0 = context, 1 = left-to-right, 2 = right-to-left. */
	readingOrder?: number;
	justifyLastLine?: boolean;
	relativeIndent?: number;
}

export interface ProtectionSpec {
	locked?: boolean;
	hidden?: boolean;
}

/** A cell format record (`<xf>`) in `cellXfs` or `cellStyleXfs`. */
export interface CellXf {
	numFmtId: number;
	fontId: number;
	fillId: number;
	borderId: number;
	/** Parent cell-style xf (index into `cellStyleXfs`); only for `cellXfs` entries. */
	xfId?: number;
	alignment?: AlignmentSpec;
	protection?: ProtectionSpec;
	applyNumberFormat?: boolean;
	applyFont?: boolean;
	applyFill?: boolean;
	applyBorder?: boolean;
	applyAlignment?: boolean;
	applyProtection?: boolean;
	/** The cell value was entered with a leading apostrophe. */
	quotePrefix?: boolean;
	pivotButton?: boolean;
}

/** A named cell style (`<cellStyle>`), e.g. "Normal", "Good", "Heading 1". */
export interface NamedCellStyle {
	name: string;
	/** Index into `cellStyleXfs`. */
	xfId: number;
	builtinId?: number;
	iLevel?: number;
	hidden?: boolean;
	customBuiltin?: boolean;
}

/** A differential format (`<dxf>`) used by conditional formatting and tables. */
export interface DifferentialFormat {
	font?: FontSpec;
	numFmt?: { id: number; code: string };
	fill?: FillSpec;
	alignment?: AlignmentSpec;
	border?: BorderSpec;
	protection?: ProtectionSpec;
}

/** One element of a custom table style (`<tableStyleElement>`). */
export interface TableStyleElement {
	type: string;
	dxfId: number;
	size?: number;
}

export interface CustomTableStyle {
	name: string;
	pivot?: boolean;
	table?: boolean;
	elements: TableStyleElement[];
}

/** Original XML for each collection entry, kept so unmodified entries are re-emitted verbatim. */
export interface StyleSheetRaw {
	/** The whole `<styleSheet>` root (used to preserve elements the model does not cover). */
	root?: XmlElement;
	fonts?: XmlElement[];
	fills?: XmlElement[];
	borders?: XmlElement[];
	cellStyleXfs?: XmlElement[];
	cellXfs?: XmlElement[];
	cellStyles?: XmlElement[];
	dxfs?: XmlElement[];
}

export interface StyleSheet {
	/** Custom number formats keyed by id (ids >= 164 by convention). Built-ins are not listed. */
	numFmts: Record<number, string>;
	fonts: FontSpec[];
	fills: FillSpec[];
	borders: BorderSpec[];
	cellStyleXfs: CellXf[];
	cellXfs: CellXf[];
	cellStyles: NamedCellStyle[];
	dxfs: DifferentialFormat[];
	/** Custom indexed palette (ARGB), when the file overrides the default 64-color palette. */
	indexedColors?: string[];
	defaultTableStyle?: string;
	defaultPivotStyle?: string;
	customTableStyles?: CustomTableStyle[];
	/** Round-trip data from the source file; absent for new workbooks. */
	raw?: StyleSheetRaw;
}

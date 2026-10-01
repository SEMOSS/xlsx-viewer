import type {
	AlignmentSpec,
	BorderEdge,
	BorderSpec,
	BorderStyle,
	CellXf,
	ColorSpec,
	DifferentialFormat,
	FillSpec,
	FontSpec,
	GradientStop,
	HorizontalAlignment,
	PatternType,
	ProtectionSpec,
	UnderlineStyle,
	VerticalAlignment,
} from '../types/style';
import type { XmlElement } from '../xml/types';
import {
	bool,
	childNum,
	childVal,
	compact,
	flag,
	int,
	kid,
	kids,
	num,
	oneOf,
	str,
} from './xml-utils';

/** Parsers for the building blocks of `styles.xml` (also used for rich-text runs and dxfs). */

/** `CT_Color`: `rgb`, `theme`, `tint`, `indexed`, `auto`. */
export function parseColor(el: XmlElement | undefined): ColorSpec | undefined {
	if (!el) {
		return undefined;
	}
	const color: ColorSpec = {};
	const rgb = str(el, 'rgb');
	if (rgb) {
		color.rgb = normalizeArgb(rgb);
	}
	const theme = int(el, 'theme');
	if (theme !== undefined) {
		color.theme = theme;
	}
	const indexed = int(el, 'indexed');
	if (indexed !== undefined) {
		color.indexed = indexed;
	}
	if (bool(el, 'auto')) {
		color.auto = true;
	}
	const tint = num(el, 'tint');
	if (tint !== undefined && tint !== 0) {
		color.tint = tint;
	}
	return Object.keys(color).length > 0 ? color : undefined;
}

/** Upper-case ARGB; 6-digit RGB gets an opaque alpha. */
export function normalizeArgb(rgb: string): string {
	const hex = rgb.trim().replace(/^#/, '').toUpperCase();
	return hex.length === 6 ? `FF${hex}` : hex;
}

const UNDERLINES: readonly UnderlineStyle[] = [
	'single',
	'double',
	'singleAccounting',
	'doubleAccounting',
	'none',
];

/** `<font>` in styles.xml or `<rPr>` of a rich-text run (`rFont` is the run's font name). */
export function parseFont(el: XmlElement | undefined): FontSpec {
	const font: FontSpec = {};
	if (!el) {
		return font;
	}
	const name = childVal(el, 'name') ?? childVal(el, 'rFont');
	if (name !== undefined) {
		font.name = name;
	}
	const size = childNum(el, 'sz');
	if (size !== undefined) {
		font.size = size;
	}
	const bold = flag(el, 'b');
	if (bold !== undefined) {
		font.bold = bold;
	}
	const italic = flag(el, 'i');
	if (italic !== undefined) {
		font.italic = italic;
	}
	const u = kid(el, 'u');
	if (u) {
		font.underline = oneOf(u.attrs.val, UNDERLINES) ?? 'single';
	}
	const strike = flag(el, 'strike');
	if (strike !== undefined) {
		font.strike = strike;
	}
	const color = parseColor(kid(el, 'color'));
	if (color) {
		font.color = color;
	}
	const va = oneOf(childVal(el, 'vertAlign'), ['superscript', 'subscript', 'baseline'] as const);
	if (va) {
		font.vertAlign = va;
	}
	const family = childNum(el, 'family');
	if (family !== undefined) {
		font.family = family;
	}
	const charset = childNum(el, 'charset');
	if (charset !== undefined) {
		font.charset = charset;
	}
	const scheme = oneOf(childVal(el, 'scheme'), ['major', 'minor', 'none'] as const);
	if (scheme) {
		font.scheme = scheme;
	}
	for (const key of ['outline', 'shadow', 'condense', 'extend'] as const) {
		const v = flag(el, key);
		if (v !== undefined) {
			font[key] = v;
		}
	}
	return font;
}

const PATTERNS: readonly PatternType[] = [
	'none',
	'solid',
	'mediumGray',
	'darkGray',
	'lightGray',
	'darkHorizontal',
	'darkVertical',
	'darkDown',
	'darkUp',
	'darkGrid',
	'darkTrellis',
	'lightHorizontal',
	'lightVertical',
	'lightDown',
	'lightUp',
	'lightGrid',
	'lightTrellis',
	'gray125',
	'gray0625',
];

/**
 * `<fill>`. In a differential format a `patternFill` without `patternType`
 * but with colors means a solid fill (Excel paints the dxf `bgColor`); in
 * `styles.fills` a missing pattern type is `none`, as the schema says.
 */
export function parseFill(el: XmlElement | undefined, differential = false): FillSpec {
	const gradient = kid(el, 'gradientFill');
	if (gradient) {
		const stops: GradientStop[] = [];
		for (const stop of kids(gradient, 'stop')) {
			const color = parseColor(kid(stop, 'color'));
			stops.push({ position: num(stop, 'position') ?? 0, color: color ?? {} });
		}
		return compact({
			type: 'gradient',
			gradientType: str(gradient, 'type') === 'path' ? 'path' : 'linear',
			degree: num(gradient, 'degree'),
			left: num(gradient, 'left'),
			right: num(gradient, 'right'),
			top: num(gradient, 'top'),
			bottom: num(gradient, 'bottom'),
			stops,
		});
	}
	const pattern = kid(el, 'patternFill');
	const fgColor = parseColor(kid(pattern, 'fgColor'));
	const bgColor = parseColor(kid(pattern, 'bgColor'));
	let patternType = oneOf(str(pattern, 'patternType'), PATTERNS);
	if (!patternType) {
		patternType = differential && (fgColor || bgColor) ? 'solid' : 'none';
	}
	return compact({ type: 'pattern', patternType, fgColor, bgColor });
}

const BORDER_STYLES: readonly BorderStyle[] = [
	'none',
	'thin',
	'medium',
	'dashed',
	'dotted',
	'thick',
	'double',
	'hair',
	'mediumDashed',
	'dashDot',
	'mediumDashDot',
	'dashDotDot',
	'mediumDashDotDot',
	'slantDashDot',
];

function parseEdge(el: XmlElement | undefined): BorderEdge | undefined {
	const style = oneOf(str(el, 'style'), BORDER_STYLES);
	if (!el || !style || style === 'none') {
		return undefined;
	}
	const color = parseColor(kid(el, 'color'));
	return color ? { style, color } : { style };
}

/** `<border>`; `start` / `end` are accepted as synonyms of `left` / `right`. */
export function parseBorder(el: XmlElement | undefined): BorderSpec {
	if (!el) {
		return {};
	}
	return compact({
		left: parseEdge(kid(el, 'left') ?? kid(el, 'start')),
		right: parseEdge(kid(el, 'right') ?? kid(el, 'end')),
		top: parseEdge(kid(el, 'top')),
		bottom: parseEdge(kid(el, 'bottom')),
		diagonal: parseEdge(kid(el, 'diagonal')),
		diagonalUp: bool(el, 'diagonalUp') || undefined,
		diagonalDown: bool(el, 'diagonalDown') || undefined,
		vertical: parseEdge(kid(el, 'vertical')),
		horizontal: parseEdge(kid(el, 'horizontal')),
		outline: bool(el, 'outline'),
	});
}

const H_ALIGN: readonly HorizontalAlignment[] = [
	'general',
	'left',
	'center',
	'right',
	'fill',
	'justify',
	'centerContinuous',
	'distributed',
];
const V_ALIGN: readonly VerticalAlignment[] = ['top', 'center', 'bottom', 'justify', 'distributed'];

export function parseAlignment(el: XmlElement | undefined): AlignmentSpec | undefined {
	if (!el) {
		return undefined;
	}
	const a = compact<AlignmentSpec>({
		horizontal: oneOf(str(el, 'horizontal'), H_ALIGN),
		vertical: oneOf(str(el, 'vertical'), V_ALIGN),
		wrapText: bool(el, 'wrapText'),
		shrinkToFit: bool(el, 'shrinkToFit'),
		indent: int(el, 'indent'),
		textRotation: int(el, 'textRotation'),
		readingOrder: int(el, 'readingOrder'),
		justifyLastLine: bool(el, 'justifyLastLine'),
		relativeIndent: int(el, 'relativeIndent'),
	});
	return Object.keys(a).length > 0 ? a : undefined;
}

export function parseProtection(el: XmlElement | undefined): ProtectionSpec | undefined {
	if (!el) {
		return undefined;
	}
	const p = compact<ProtectionSpec>({ locked: bool(el, 'locked'), hidden: bool(el, 'hidden') });
	return Object.keys(p).length > 0 ? p : undefined;
}

/** `<xf>` of `cellXfs` (with `xfId`) or `cellStyleXfs` (without). */
export function parseXf(el: XmlElement, cellXf: boolean): CellXf {
	const xf: CellXf = {
		numFmtId: int(el, 'numFmtId') ?? 0,
		fontId: int(el, 'fontId') ?? 0,
		fillId: int(el, 'fillId') ?? 0,
		borderId: int(el, 'borderId') ?? 0,
	};
	if (cellXf) {
		xf.xfId = int(el, 'xfId') ?? 0;
	}
	const alignment = parseAlignment(kid(el, 'alignment'));
	if (alignment) {
		xf.alignment = alignment;
	}
	const protection = parseProtection(kid(el, 'protection'));
	if (protection) {
		xf.protection = protection;
	}
	for (const key of [
		'applyNumberFormat',
		'applyFont',
		'applyFill',
		'applyBorder',
		'applyAlignment',
		'applyProtection',
		'quotePrefix',
		'pivotButton',
	] as const) {
		const v = bool(el, key);
		if (v !== undefined) {
			xf[key] = v;
		}
	}
	return xf;
}

/** `<dxf>` (also the `x14:dxf` of Excel 2010 conditional-formatting rules). */
export function parseDxf(el: XmlElement | undefined): DifferentialFormat {
	const dxf: DifferentialFormat = {};
	if (!el) {
		return dxf;
	}
	const font = kid(el, 'font');
	if (font) {
		dxf.font = parseFont(font);
	}
	const numFmt = kid(el, 'numFmt');
	if (numFmt) {
		dxf.numFmt = { id: int(numFmt, 'numFmtId') ?? 0, code: str(numFmt, 'formatCode') ?? 'General' };
	}
	const fill = kid(el, 'fill');
	if (fill) {
		dxf.fill = parseFill(fill, true);
	}
	const alignment = parseAlignment(kid(el, 'alignment'));
	if (alignment) {
		dxf.alignment = alignment;
	}
	const border = kid(el, 'border');
	if (border) {
		dxf.border = parseBorder(border);
	}
	const protection = parseProtection(kid(el, 'protection'));
	if (protection) {
		dxf.protection = protection;
	}
	return dxf;
}

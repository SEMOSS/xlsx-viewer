/** Serializers for individual styles.xml entries (fonts, fills, borders, xfs, dxfs, ...). */
import type {
	AlignmentSpec,
	BorderEdge,
	BorderSpec,
	CellXf,
	CustomTableStyle,
	DifferentialFormat,
	FillSpec,
	NamedCellStyle,
	ProtectionSpec,
} from '../types/style';
import { fontXml } from './rich-text';
import { attrString, colorTag, flag, tag } from './xml-utils';

export { fontXml };

/** `<numFmt>`. */
export function numFmtXml(id: number, code: string): string {
	return tag('numFmt', { numFmtId: id, formatCode: code });
}

function clamp01(n: number): number {
	return Math.max(0, Math.min(1, n));
}

/**
 * `<fill>`. In differential formats a solid fill is written with both
 * colors set so it renders the same whichever color Excel reads.
 */
export function fillXml(fill: FillSpec, dxf = false): string {
	if (fill.type === 'gradient') {
		const attrs = {
			type: fill.gradientType === 'path' ? 'path' : undefined,
			degree: fill.gradientType === 'path' ? undefined : fill.degree,
			left: fill.left,
			right: fill.right,
			top: fill.top,
			bottom: fill.bottom,
		};
		const stops = fill.stops
			.map((s) => `<stop position="${clamp01(s.position)}">${colorTag('color', s.color)}</stop>`)
			.join('');
		return `<fill><gradientFill${attrString(attrs)}>${stops}</gradientFill></fill>`;
	}
	let fg = fill.fgColor;
	let bg = fill.bgColor;
	if (dxf && fill.patternType === 'solid') {
		fg ??= bg;
		bg ??= fg;
	}
	const inner = colorTag('fgColor', fg) + colorTag('bgColor', bg);
	return `<fill>${tag('patternFill', { patternType: fill.patternType }, inner)}</fill>`;
}

function edgeXml(name: string, edge: BorderEdge | undefined, always: boolean): string {
	if (!edge || edge.style === 'none') {
		return always ? `<${name}/>` : '';
	}
	return tag(name, { style: edge.style }, colorTag('color', edge.color));
}

/**
 * `<border>`. Cell borders always list the five standard edges (as Excel
 * does); differential borders only the edges that are set.
 */
export function borderXml(b: BorderSpec, dxf = false): string {
	const attrs = {
		diagonalUp: flag(b.diagonalUp),
		diagonalDown: flag(b.diagonalDown),
		outline: b.outline === false ? '0' : undefined,
	};
	const always = !dxf;
	const inner =
		edgeXml('left', b.left, always) +
		edgeXml('right', b.right, always) +
		edgeXml('top', b.top, always) +
		edgeXml('bottom', b.bottom, always) +
		edgeXml('diagonal', b.diagonal, always) +
		edgeXml('vertical', b.vertical, false) +
		edgeXml('horizontal', b.horizontal, false);
	return tag('border', attrs, inner);
}

/** `<alignment>`. */
export function alignmentXml(a: AlignmentSpec): string {
	return tag('alignment', {
		horizontal: a.horizontal,
		vertical: a.vertical,
		textRotation: a.textRotation || undefined,
		wrapText: a.wrapText,
		indent: a.indent || undefined,
		relativeIndent: a.relativeIndent || undefined,
		justifyLastLine: a.justifyLastLine,
		shrinkToFit: a.shrinkToFit,
		readingOrder: a.readingOrder || undefined,
	});
}

/** `<protection>`. */
export function protectionXml(p: ProtectionSpec): string {
	return tag('protection', { locked: p.locked, hidden: p.hidden });
}

function hasKeys(o: object | undefined): boolean {
	if (!o) {
		return false;
	}
	for (const k in o) {
		if ((o as Record<string, unknown>)[k] !== undefined) {
			return true;
		}
	}
	return false;
}

/** `<xf>` for `cellXfs` (with `xfId`) or `cellStyleXfs`. */
export function xfXml(xf: CellXf, cellXf: boolean): string {
	const attrs = {
		numFmtId: xf.numFmtId ?? 0,
		fontId: xf.fontId ?? 0,
		fillId: xf.fillId ?? 0,
		borderId: xf.borderId ?? 0,
		xfId: cellXf ? (xf.xfId ?? 0) : undefined,
		quotePrefix: flag(xf.quotePrefix),
		pivotButton: flag(xf.pivotButton),
		applyNumberFormat: xf.applyNumberFormat,
		applyFont: xf.applyFont,
		applyFill: xf.applyFill,
		applyBorder: xf.applyBorder,
		applyAlignment: xf.applyAlignment,
		applyProtection: xf.applyProtection,
	};
	const inner =
		(hasKeys(xf.alignment) ? alignmentXml(xf.alignment!) : '') +
		(hasKeys(xf.protection) ? protectionXml(xf.protection!) : '');
	return tag('xf', attrs, inner);
}

/** `<cellStyle>`. */
export function cellStyleXml(s: NamedCellStyle): string {
	return tag('cellStyle', {
		name: s.name,
		xfId: s.xfId,
		builtinId: s.builtinId,
		iLevel: s.iLevel,
		hidden: flag(s.hidden),
		customBuiltin: flag(s.customBuiltin),
	});
}

/** Children of a `<dxf>` in CT_Dxf order (also used for inline `x14:dxf`). */
export function dxfInnerXml(d: DifferentialFormat): string {
	return (
		(d.font ? fontXml(d.font) : '') +
		(d.numFmt ? numFmtXml(d.numFmt.id, d.numFmt.code) : '') +
		(d.fill ? fillXml(d.fill, true) : '') +
		(hasKeys(d.alignment) ? alignmentXml(d.alignment!) : '') +
		(hasKeys(d.protection) ? protectionXml(d.protection!) : '') +
		(d.border ? borderXml(d.border, true) : '')
	);
}

/** `<dxf>`. */
export function dxfXml(d: DifferentialFormat): string {
	const inner = dxfInnerXml(d);
	return inner ? `<dxf>${inner}</dxf>` : '<dxf/>';
}

/** `<tableStyle>` (a custom table / pivot style). */
export function tableStyleXml(s: CustomTableStyle): string {
	const inner = s.elements
		.map((e) => tag('tableStyleElement', { type: e.type, size: e.size, dxfId: e.dxfId }))
		.join('');
	return tag(
		'tableStyle',
		{
			name: s.name,
			pivot: s.pivot === false ? '0' : undefined,
			table: s.table === false ? '0' : undefined,
			count: s.elements.length,
		},
		inner,
	);
}

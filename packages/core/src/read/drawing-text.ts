import type {
	DrawingColor,
	DrawingParagraph,
	DrawingTextBody,
	DrawingTextStyle,
} from '../types/drawing';
import { isElement, textContent } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { coordinate, drawingNumber, findColor } from './drawingml';
import { kid, kids, lname, oneOf, str } from './xml-utils';

/** DrawingML text bodies (`txBody`) of shapes and text boxes. */

/** Run properties (`a:rPr`, `a:defRPr`, `a:endParaRPr`). */
export function parseRunStyle(el: XmlElement | undefined): DrawingTextStyle | undefined {
	if (!el) {
		return undefined;
	}
	const style: DrawingTextStyle = {};
	const sz = drawingNumber(el.attrs.sz);
	if (sz !== undefined) {
		style.size = sz / 100;
	}
	if (el.attrs.b !== undefined) {
		style.bold = el.attrs.b === '1' || el.attrs.b === 'true';
	}
	if (el.attrs.i !== undefined) {
		style.italic = el.attrs.i === '1' || el.attrs.i === 'true';
	}
	if (el.attrs.u !== undefined) {
		style.underline = el.attrs.u !== 'none';
	}
	if (el.attrs.strike !== undefined) {
		style.strike = el.attrs.strike !== 'noStrike';
	}
	const fill = kid(el, 'solidFill');
	const color = fill ? findColor(fill) : undefined;
	if (color) {
		style.color = color;
	}
	const font = str(kid(el, 'latin'), 'typeface');
	if (font) {
		style.font = font;
	}
	return Object.keys(style).length > 0 ? style : undefined;
}

function mergeStyles(
	base: DrawingTextStyle | undefined,
	over: DrawingTextStyle | undefined,
): DrawingTextStyle | undefined {
	if (!base) {
		return over;
	}
	return over ? { ...base, ...over } : { ...base };
}

const ALIGNS = ['l', 'ctr', 'r', 'just', 'dist'] as const;

/**
 * `txBody`: body properties, paragraphs and runs. Line breaks become `\n`
 * runs; an empty paragraph keeps one empty run carrying its end-of-paragraph
 * style so renderers know the line height. `defaultColor` (from the shape
 * style's `fontRef`) applies to runs without an explicit color.
 */
export function parseTextBody(
	el: XmlElement | undefined,
	defaultColor?: DrawingColor,
): DrawingTextBody | undefined {
	if (!el) {
		return undefined;
	}
	const body: DrawingTextBody = { paragraphs: [] };
	const bodyPr = kid(el, 'bodyPr');
	if (bodyPr) {
		const anchor = oneOf(str(bodyPr, 'anchor'), ['t', 'ctr', 'b'] as const);
		if (anchor) {
			body.anchor = anchor;
		}
		const insets: NonNullable<DrawingTextBody['insets']> = {};
		for (const [attr, key] of [
			['lIns', 'l'],
			['tIns', 't'],
			['rIns', 'r'],
			['bIns', 'b'],
		] as const) {
			const v = coordinate(bodyPr.attrs[attr]);
			if (v !== undefined) {
				insets[key] = v;
			}
		}
		if (Object.keys(insets).length > 0) {
			body.insets = insets;
		}
		const wrap = str(bodyPr, 'wrap');
		if (wrap) {
			body.wrap = wrap !== 'none';
		}
		const vert = str(bodyPr, 'vert');
		if (vert) {
			body.vertical = vert;
		}
	}
	const base: DrawingTextStyle | undefined = defaultColor ? { color: defaultColor } : undefined;
	for (const p of kids(el, 'p')) {
		const pPr = kid(p, 'pPr');
		const paraStyle = mergeStyles(base, parseRunStyle(kid(pPr, 'defRPr')));
		const para: DrawingParagraph = { runs: [] };
		const align = oneOf(str(pPr, 'algn'), ALIGNS);
		if (align) {
			para.align = align;
		}
		for (const c of p.children) {
			if (!isElement(c)) {
				continue;
			}
			const name = lname(c);
			if (name === 'r' || name === 'fld') {
				const style = mergeStyles(paraStyle, parseRunStyle(kid(c, 'rPr')));
				const run: DrawingParagraph['runs'][number] = { text: textContent(kid(c, 't')) };
				if (style) {
					run.style = style;
				}
				para.runs.push(run);
			} else if (name === 'br') {
				const style = mergeStyles(paraStyle, parseRunStyle(kid(c, 'rPr')));
				para.runs.push(style ? { text: '\n', style } : { text: '\n' });
			}
		}
		if (para.runs.length === 0) {
			const end = mergeStyles(paraStyle, parseRunStyle(kid(p, 'endParaRPr')));
			if (end) {
				para.runs.push({ text: '', style: end });
			}
		}
		body.paragraphs.push(para);
	}
	return body;
}

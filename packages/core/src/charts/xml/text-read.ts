/** Readers for DrawingML text (run properties, text bodies, rich text) and manual layouts. */
import type { ChartManualLayout } from '../../types/chart';
import type { DrawingParagraph, DrawingTextStyle } from '../../types/drawing';
import { isElement, localName, textContent } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';
import { kid, kids, numAttr, numOf, parseBool, readFill, valOf } from './drawingml-read';

/** Read run properties (`a:defRPr`, `a:rPr`, `a:endParaRPr`). */
export function readTextStyle(rPr: XmlElement | undefined): DrawingTextStyle | undefined {
	if (!rPr) {
		return undefined;
	}
	const style: DrawingTextStyle = {};
	const sz = numAttr(rPr, 'sz');
	if (sz !== undefined) {
		style.size = sz / 100;
	}
	if (rPr.attrs.b !== undefined) {
		style.bold = parseBool(rPr.attrs.b);
	}
	if (rPr.attrs.i !== undefined) {
		style.italic = parseBool(rPr.attrs.i);
	}
	if (rPr.attrs.u !== undefined) {
		style.underline = rPr.attrs.u !== 'none';
	}
	if (rPr.attrs.strike !== undefined) {
		style.strike = rPr.attrs.strike !== 'noStrike';
	}
	const fill = readFill(rPr);
	if (fill?.type === 'solid') {
		style.color = fill.color;
	} else if (fill?.type === 'gradient' && fill.stops.length) {
		style.color = fill.stops[0].color;
	}
	const latin = kid(rPr, 'latin')?.attrs.typeface;
	if (latin) {
		style.font = latin;
	}
	return Object.keys(style).length ? style : undefined;
}

/** Text rotation from `a:bodyPr` in degrees; `undefined` for automatic. */
export function readBodyRotation(bodyPr: XmlElement | undefined): number | undefined {
	if (!bodyPr) {
		return undefined;
	}
	const rot = numAttr(bodyPr, 'rot');
	// -60000000 is Excel's "automatic rotation" sentinel.
	if (rot !== undefined && rot !== -60000000) {
		return rot / 60000;
	}
	const vert = bodyPr.attrs.vert;
	if (vert === 'vert' || vert === 'eaVert' || vert === 'wordArtVert') {
		return 90;
	}
	if (vert === 'vert270') {
		return -90;
	}
	return undefined;
}

/** Default style and rotation of a `c:txPr`. */
export function readTxPr(txPr: XmlElement | undefined): {
	style?: DrawingTextStyle;
	rotation?: number;
} {
	if (!txPr) {
		return {};
	}
	const out: { style?: DrawingTextStyle; rotation?: number } = {};
	const p = kid(txPr, 'p');
	const style = readTextStyle(kid(kid(p, 'pPr'), 'defRPr'));
	if (style) {
		out.style = style;
	}
	const rotation = readBodyRotation(kid(txPr, 'bodyPr'));
	if (rotation !== undefined) {
		out.rotation = rotation;
	}
	return out;
}

const ALIGNS = new Set(['l', 'ctr', 'r', 'just', 'dist']);

/** Read a `c:rich` text body into paragraphs, plain text and its default style. */
export function readRichText(rich: XmlElement): {
	text: string;
	paragraphs: DrawingParagraph[];
	style?: DrawingTextStyle;
	rotation?: number;
} {
	const paragraphs: DrawingParagraph[] = [];
	let style: DrawingTextStyle | undefined;
	for (const p of kids(rich, 'p')) {
		const pPr = kid(p, 'pPr');
		const para: DrawingParagraph = { runs: [] };
		const algn = pPr?.attrs.algn;
		if (algn && ALIGNS.has(algn)) {
			para.align = algn as DrawingParagraph['align'];
		}
		if (!style) {
			style = readTextStyle(kid(pPr, 'defRPr'));
		}
		for (const r of p.children) {
			if (!isElement(r)) {
				continue;
			}
			const local = localName(r.name);
			if (local === 'r' || local === 'fld') {
				const runStyle = readTextStyle(kid(r, 'rPr'));
				const run: { text: string; style?: DrawingTextStyle } = { text: textContent(kid(r, 't')) };
				if (runStyle) {
					run.style = runStyle;
				}
				para.runs.push(run);
			} else if (local === 'br') {
				para.runs.push({ text: '\n' });
			}
		}
		paragraphs.push(para);
	}
	const text = paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
	const out: {
		text: string;
		paragraphs: DrawingParagraph[];
		style?: DrawingTextStyle;
		rotation?: number;
	} = {
		text,
		paragraphs,
	};
	if (style) {
		out.style = style;
	}
	const rotation = readBodyRotation(kid(rich, 'bodyPr'));
	if (rotation !== undefined) {
		out.rotation = rotation;
	}
	return out;
}

/** `c:layout/c:manualLayout`, when present. */
export function readManualLayout(layout: XmlElement | undefined): ChartManualLayout | undefined {
	const ml = kid(layout, 'manualLayout');
	if (!ml) {
		return undefined;
	}
	const out: ChartManualLayout = {};
	const target = valOf(ml, 'layoutTarget');
	if (target === 'inner' || target === 'outer') {
		out.layoutTarget = target;
	}
	for (const mode of ['xMode', 'yMode'] as const) {
		const v = valOf(ml, mode);
		if (v === 'edge' || v === 'factor') {
			out[mode] = v;
		}
	}
	for (const key of ['x', 'y', 'w', 'h'] as const) {
		const v = numOf(ml, key);
		if (v !== undefined) {
			out[key] = v;
		}
	}
	return Object.keys(out).length ? out : undefined;
}

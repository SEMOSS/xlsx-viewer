/** Builders for the DrawingML fragments of a chart part (colors, fills, outlines, text). */
import type { ChartManualLayout, ChartText } from '../../types/chart';
import type {
	DrawingColor,
	DrawingParagraph,
	DrawingTextStyle,
	ShapeFill,
	ShapeLine,
} from '../../types/drawing';
import { el } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';

/** `<name val="..."/>`. */
export function valEl(name: string, val: string | number | boolean): XmlElement {
	return el(name, { val: typeof val === 'boolean' ? (val ? '1' : '0') : String(val) });
}

/** Number text the way Excel writes it (`1E-3`, not `1e-3`). */
export function formatXmlNumber(n: number): string {
	return String(n).replace('e', 'E');
}

export function writeColor(color: DrawingColor): XmlElement {
	const transforms = (color.transforms ?? []).map((t) =>
		el(`a:${t.name}`, t.val === undefined ? {} : { val: Math.round(t.val) }),
	);
	switch (color.kind) {
		case 'scheme':
			return el('a:schemeClr', { val: color.value }, transforms);
		case 'preset':
			return el('a:prstClr', { val: color.value }, transforms);
		case 'sys':
			return /^[0-9a-fA-F]{6}$/.test(color.value)
				? el('a:srgbClr', { val: color.value.toUpperCase() }, transforms)
				: el('a:sysClr', { val: color.value }, transforms);
		case 'hsl': {
			const [hue, sat, lum] = color.value.split(',');
			return el('a:hslClr', { hue, sat, lum }, transforms);
		}
		case 'scrgb': {
			const [r, g, b] = color.value.split(',');
			return el('a:scrgbClr', { r, g, b }, transforms);
		}
		default:
			return el('a:srgbClr', { val: color.value.replace(/^#/, '').toUpperCase() }, transforms);
	}
}

/** Fill element, or `undefined` for fills that cannot be regenerated (pictures). */
export function writeFill(fill: ShapeFill | undefined): XmlElement | undefined {
	if (!fill) {
		return undefined;
	}
	switch (fill.type) {
		case 'none':
			return el('a:noFill');
		case 'solid':
			return el('a:solidFill', {}, [writeColor(fill.color)]);
		case 'gradient': {
			const stops = fill.stops.map((s) =>
				el('a:gs', { pos: Math.round((s.pos > 1 ? s.pos / 100000 : s.pos) * 100000) }, [
					writeColor(s.color),
				]),
			);
			const shade =
				fill.path !== undefined
					? el('a:path', { path: fill.path }, [
							el('a:fillToRect', { l: 50000, t: 50000, r: 50000, b: 50000 }),
						])
					: el('a:lin', { ang: Math.round((fill.angle ?? 0) * 60000), scaled: 0 });
			return el('a:gradFill', { rotWithShape: 1 }, [el('a:gsLst', {}, stops), shade]);
		}
		case 'pattern':
			return el('a:pattFill', { prst: fill.preset }, [
				fill.fg ? el('a:fgClr', {}, [writeColor(fill.fg)]) : undefined,
				fill.bg ? el('a:bgClr', {}, [writeColor(fill.bg)]) : undefined,
			]);
		default:
			return undefined;
	}
}

export function writeLine(line: ShapeLine): XmlElement {
	const fill = line.fill?.type === 'picture' ? undefined : writeFill(line.fill);
	const end = (name: string, spec: ShapeLine['headEnd']) =>
		spec ? el(name, { type: spec.type, w: spec.w, len: spec.len }) : undefined;
	return el(
		'a:ln',
		{ w: line.width === undefined ? undefined : Math.round(line.width), cap: line.cap },
		[
			fill,
			line.dash ? valEl('a:prstDash', line.dash) : undefined,
			end('a:headEnd', line.headEnd),
			end('a:tailEnd', line.tailEnd),
		],
	);
}

/** `c:spPr` with a fill and/or outline, or `undefined` when both are absent. */
export function writeShapeProps(
	fill: ShapeFill | undefined,
	line: ShapeLine | undefined,
	name = 'c:spPr',
): XmlElement | undefined {
	const f = writeFill(fill);
	if (!f && !line) {
		return undefined;
	}
	return el(name, {}, [f, line ? writeLine(line) : undefined]);
}

/** Run properties (`a:defRPr` or `a:rPr`). */
export function writeRunProps(name: string, style: DrawingTextStyle | undefined): XmlElement {
	const s = style ?? {};
	const attrs: Record<string, string | number | undefined> = {};
	if (name === 'a:rPr') {
		attrs.lang = 'en-US';
	}
	attrs.sz = s.size === undefined ? undefined : Math.round(s.size * 100);
	attrs.b = s.bold === undefined ? undefined : s.bold ? 1 : 0;
	attrs.i = s.italic === undefined ? undefined : s.italic ? 1 : 0;
	attrs.u = s.underline === undefined ? undefined : s.underline ? 'sng' : 'none';
	attrs.strike = s.strike === undefined ? undefined : s.strike ? 'sngStrike' : 'noStrike';
	return el(name, attrs, [
		s.color ? el('a:solidFill', {}, [writeColor(s.color)]) : undefined,
		s.font ? el('a:latin', { typeface: s.font }) : undefined,
	]);
}

function bodyPr(rotation: number | undefined): XmlElement {
	return rotation === undefined
		? el('a:bodyPr')
		: el('a:bodyPr', { rot: Math.round(rotation * 60000), vert: 'horz' });
}

/** `c:txPr` carrying a default text style and rotation, or `undefined` when there is neither. */
export function writeTextProps(
	style: DrawingTextStyle | undefined,
	rotation?: number,
): XmlElement | undefined {
	if (!style && rotation === undefined) {
		return undefined;
	}
	return el('c:txPr', {}, [
		bodyPr(rotation),
		el('a:lstStyle'),
		el('a:p', {}, [
			el('a:pPr', {}, [writeRunProps('a:defRPr', style)]),
			el('a:endParaRPr', { lang: 'en-US' }),
		]),
	]);
}

function textParagraphs(text: ChartText): DrawingParagraph[] {
	if (text.paragraphs?.length) {
		return text.paragraphs;
	}
	return (text.text ?? '').split('\n').map((line) => ({ runs: [{ text: line }] }));
}

/** `c:rich` text body for a title. */
export function writeRichText(text: ChartText): XmlElement {
	const paras = textParagraphs(text).map((p) => {
		const runs: XmlElement[] = [];
		for (const r of p.runs) {
			if (r.text === '\n') {
				runs.push(el('a:br', {}, [writeRunProps('a:rPr', r.style)]));
				continue;
			}
			const t = el('a:t', /^\s|\s$/.test(r.text) ? { 'xml:space': 'preserve' } : {}, [r.text]);
			runs.push(el('a:r', {}, [writeRunProps('a:rPr', r.style), t]));
		}
		return el('a:p', {}, [
			el('a:pPr', p.align ? { algn: p.align } : {}, [writeRunProps('a:defRPr', text.style)]),
			...runs,
		]);
	});
	return el('c:rich', {}, [bodyPr(text.rotation), el('a:lstStyle'), ...paras]);
}

/** `c:layout`, with a manual layout when given (an empty `c:layout` means automatic). */
export function writeLayout(layout: ChartManualLayout | undefined): XmlElement {
	if (!layout) {
		return el('c:layout');
	}
	const num = (name: string, v: number | undefined) =>
		v === undefined ? undefined : valEl(name, formatXmlNumber(v));
	return el('c:layout', {}, [
		el('c:manualLayout', {}, [
			layout.layoutTarget ? valEl('c:layoutTarget', layout.layoutTarget) : undefined,
			layout.xMode ? valEl('c:xMode', layout.xMode) : undefined,
			layout.yMode ? valEl('c:yMode', layout.yMode) : undefined,
			num('c:x', layout.x),
			num('c:y', layout.y),
			num('c:w', layout.w),
			num('c:h', layout.h),
		]),
	]);
}

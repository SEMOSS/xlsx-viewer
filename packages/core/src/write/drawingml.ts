/**
 * DrawingML primitives for regenerated drawing objects: colors, fills,
 * outlines and text bodies (prefix `a:`).
 */
import type {
	DrawingColor,
	DrawingParagraph,
	DrawingTextBody,
	DrawingTextStyle,
	ShapeFill,
	ShapeLine,
} from '../types/drawing';
import { escapeXmlText } from '../xml';
import { tag } from './xml-utils';

/** Resolves a media path to the relationship id of its image in the drawing part. */
export type BlipRelResolver = (mediaPath: string) => string | undefined;

function hex6(value: string): string {
	const v = value.replace(/^#/, '').toUpperCase();
	return /^[0-9A-F]{8}$/.test(v) ? v.slice(2) : v;
}

function transformsXml(c: DrawingColor): string {
	return (c.transforms ?? [])
		.filter((t) => /^[A-Za-z]+$/.test(t.name))
		.map((t) =>
			t.val === undefined ? `<a:${t.name}/>` : `<a:${t.name} val="${Math.round(t.val)}"/>`,
		)
		.join('');
}

/** A DrawingML color element. */
export function colorXml(c: DrawingColor): string {
	const tr = transformsXml(c);
	switch (c.kind) {
		case 'srgb':
			return tag('a:srgbClr', { val: hex6(c.value) }, tr);
		case 'scheme':
			return tag('a:schemeClr', { val: c.value }, tr);
		case 'sys':
			return /^[0-9A-Fa-f]{6}$/.test(c.value)
				? tag('a:srgbClr', { val: hex6(c.value) }, tr)
				: tag('a:sysClr', { val: c.value }, tr);
		case 'preset':
			return tag('a:prstClr', { val: c.value }, tr);
		case 'hsl': {
			const [hue, sat, lum] = c.value.split(',').map((s) => Number(s.trim()) || 0);
			return tag('a:hslClr', { hue, sat, lum }, tr);
		}
		case 'scrgb': {
			const [r, g, b] = c.value.split(',').map((s) => Number(s.trim()) || 0);
			return tag('a:scrgbClr', { r, g, b }, tr);
		}
		default:
			return tag('a:srgbClr', { val: '000000' });
	}
}

/** A fill element (`a:noFill`, `a:solidFill`, ...). Picture fills need `blipRel`. */
export function fillXml(f: ShapeFill, blipRel?: BlipRelResolver): string {
	switch (f.type) {
		case 'none':
			return '<a:noFill/>';
		case 'solid':
			return `<a:solidFill>${colorXml(f.color)}</a:solidFill>`;
		case 'gradient': {
			const stops = f.stops
				.map((s) => {
					const pos = Math.round(Math.max(0, Math.min(1, s.pos)) * 100000);
					return `<a:gs pos="${pos}">${colorXml(s.color)}</a:gs>`;
				})
				.join('');
			const shade = f.path
				? `<a:path path="${f.path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`
				: `<a:lin ang="${Math.round((f.angle ?? 90) * 60000)}" scaled="0"/>`;
			return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`;
		}
		case 'pattern':
			return tag(
				'a:pattFill',
				{ prst: f.preset },
				(f.fg ? `<a:fgClr>${colorXml(f.fg)}</a:fgClr>` : '') +
					(f.bg ? `<a:bgClr>${colorXml(f.bg)}</a:bgClr>` : ''),
			);
		case 'picture': {
			const id = blipRel?.(f.mediaPath);
			return id
				? `<a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`
				: '<a:noFill/>';
		}
		default:
			return '';
	}
}

function endXml(name: string, end: ShapeLine['headEnd']): string {
	return end ? tag(name, { type: end.type, w: end.w, len: end.len }) : '';
}

/** An `a:ln` outline. */
export function lineXml(l: ShapeLine, blipRel?: BlipRelResolver): string {
	const inner =
		(l.fill ? fillXml(l.fill, blipRel) : '') +
		(l.dash ? tag('a:prstDash', { val: l.dash }) : '') +
		endXml('a:headEnd', l.headEnd) +
		endXml('a:tailEnd', l.tailEnd);
	return tag(
		'a:ln',
		{ w: l.width !== undefined ? Math.round(l.width) : undefined, cap: l.cap },
		inner,
	);
}

/** Run properties (`a:rPr` / `a:endParaRPr` / `a:defRPr`). */
export function runPropsXml(style: DrawingTextStyle | undefined, name = 'a:rPr'): string {
	const s = style ?? {};
	const inner =
		(s.color ? `<a:solidFill>${colorXml(s.color)}</a:solidFill>` : '') +
		(s.font ? tag('a:latin', { typeface: s.font }) : '');
	return tag(
		name,
		{
			lang: 'en-US',
			sz: s.size !== undefined ? Math.round(s.size * 100) : undefined,
			b: s.bold,
			i: s.italic,
			u: s.underline === undefined ? undefined : s.underline ? 'sng' : 'none',
			strike: s.strike === undefined ? undefined : s.strike ? 'sngStrike' : 'noStrike',
		},
		inner,
	);
}

function runsXml(p: DrawingParagraph): string {
	let out = '';
	for (const run of p.runs) {
		const pieces = run.text.split(/\r\n|\n|\r/);
		pieces.forEach((piece, i) => {
			if (i > 0) {
				out += `<a:br>${runPropsXml(run.style)}</a:br>`;
			}
			if (piece !== '' || pieces.length === 1) {
				out += `<a:r>${runPropsXml(run.style)}<a:t>${escapeXmlText(piece)}</a:t></a:r>`;
			}
		});
	}
	return out;
}

/** A text body (`xdr:txBody` for shapes); always holds at least one paragraph. */
export function textBodyXml(tb: DrawingTextBody | undefined, name = 'xdr:txBody'): string {
	const body = tb ?? { paragraphs: [] };
	const ins = body.insets ?? {};
	const bodyPr = tag('a:bodyPr', {
		vertOverflow: 'clip',
		horzOverflow: 'clip',
		vert: body.vertical,
		wrap: body.wrap === false ? 'none' : 'square',
		lIns: ins.l,
		tIns: ins.t,
		rIns: ins.r,
		bIns: ins.b,
		rtlCol: '0',
		anchor: body.anchor ?? 't',
	});
	const paragraphs = body.paragraphs.length ? body.paragraphs : [{ runs: [] }];
	const ps = paragraphs
		.map((p) => {
			const ppr = p.align ? tag('a:pPr', { algn: p.align }) : '';
			const last = p.runs[p.runs.length - 1]?.style;
			return `<a:p>${ppr}${runsXml(p)}${runPropsXml(last, 'a:endParaRPr')}</a:p>`;
		})
		.join('');
	return `<${name}>${bodyPr}<a:lstStyle/>${ps}</${name}>`;
}

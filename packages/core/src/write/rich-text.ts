/**
 * Font properties <-> XML for cell fonts (`<font>`), rich-text runs
 * (`<rPr>`, which names the face `rFont`) and differential formats.
 */
import type { RichTextRun } from '../types/cell';
import type { ColorSpec, FontSpec, UnderlineStyle } from '../types/style';
import { attrNum, child, elements, escapeXmlAttr, localName } from '../xml';
import type { XmlElement } from '../xml/types';
import { colorTag, formatXmlNumber, textElement } from './xml-utils';

function toggle(p: string, name: string, v: boolean | undefined): string {
	if (v === undefined) {
		return '';
	}
	return v ? `<${p}${name}/>` : `<${p}${name} val="0"/>`;
}

function valTag(p: string, name: string, v: string | number | undefined): string {
	if (v === undefined) {
		return '';
	}
	const text = typeof v === 'number' ? formatXmlNumber(v) : escapeXmlAttr(v);
	return `<${p}${name} val="${text}"/>`;
}

/**
 * Font property children in the order Excel writes them. `nameTag` is `name`
 * for `<font>` and `rFont` for rich-text run properties.
 */
export function fontPropsXml(font: FontSpec, nameTag: 'name' | 'rFont', p = ''): string {
	let s = '';
	s += toggle(p, 'b', font.bold);
	s += toggle(p, 'i', font.italic);
	s += toggle(p, 'strike', font.strike);
	s += toggle(p, 'condense', font.condense);
	s += toggle(p, 'extend', font.extend);
	s += toggle(p, 'outline', font.outline);
	s += toggle(p, 'shadow', font.shadow);
	if (font.underline !== undefined) {
		s += font.underline === 'single' ? `<${p}u/>` : `<${p}u val="${font.underline}"/>`;
	}
	s += valTag(p, 'vertAlign', font.vertAlign);
	s += valTag(p, 'sz', font.size);
	s += colorTag(`${p}color`, font.color);
	s += valTag(p, nameTag, font.name);
	s += valTag(p, 'family', font.family);
	s += valTag(p, 'charset', font.charset);
	s += valTag(p, 'scheme', font.scheme);
	return s;
}

/** A `<font>` element for styles.xml / dxfs. */
export function fontXml(font: FontSpec, p = ''): string {
	const inner = fontPropsXml(font, 'name', p);
	return inner ? `<${p}font>${inner}</${p}font>` : `<${p}font/>`;
}

/** The font has no properties set. */
export function isEmptyFont(font: FontSpec | undefined): boolean {
	if (!font) {
		return true;
	}
	for (const k in font) {
		if ((font as Record<string, unknown>)[k] !== undefined) {
			return false;
		}
	}
	return true;
}

/** Rich-text runs as `<r>` elements (shared strings and comments). */
export function richRunsXml(runs: readonly RichTextRun[], p = ''): string {
	let s = '';
	for (const run of runs) {
		const props = run.font ? fontPropsXml(run.font, 'rFont', p) : '';
		s += `<${p}r>${props ? `<${p}rPr>${props}</${p}rPr>` : ''}${textElement(run.text, p)}</${p}r>`;
	}
	return s;
}

// ── Parsing (used to match original shared strings) ──────────────────────

function parseColor(el: XmlElement): ColorSpec {
	const c: ColorSpec = {};
	if (el.attrs.rgb !== undefined) {
		c.rgb = el.attrs.rgb;
	}
	const theme = attrNum(el, 'theme');
	if (theme !== undefined) {
		c.theme = theme;
	}
	const tint = attrNum(el, 'tint');
	if (tint !== undefined && tint !== 0) {
		c.tint = tint;
	}
	const indexed = attrNum(el, 'indexed');
	if (indexed !== undefined) {
		c.indexed = indexed;
	}
	if (el.attrs.auto === '1' || el.attrs.auto === 'true') {
		c.auto = true;
	}
	return c;
}

function flagOf(el: XmlElement): boolean {
	const v = el.attrs.val;
	return v === undefined || v === '1' || v.toLowerCase() === 'true';
}

/** Parse `<rPr>` / `<font>` children into a FontSpec (only the properties the model knows). */
export function parseFontProps(el: XmlElement): FontSpec {
	const f: FontSpec = {};
	for (const c of elements(el)) {
		const v = c.attrs.val;
		switch (localName(c.name)) {
			case 'b':
				f.bold = flagOf(c);
				break;
			case 'i':
				f.italic = flagOf(c);
				break;
			case 'strike':
				f.strike = flagOf(c);
				break;
			case 'condense':
				f.condense = flagOf(c);
				break;
			case 'extend':
				f.extend = flagOf(c);
				break;
			case 'outline':
				f.outline = flagOf(c);
				break;
			case 'shadow':
				f.shadow = flagOf(c);
				break;
			case 'u':
				f.underline = (v ?? 'single') as UnderlineStyle;
				break;
			case 'vertAlign':
				f.vertAlign = v as FontSpec['vertAlign'];
				break;
			case 'sz':
				f.size = v !== undefined ? Number(v) : undefined;
				break;
			case 'color':
				f.color = parseColor(c);
				break;
			case 'rFont':
			case 'name':
				f.name = v;
				break;
			case 'family':
				f.family = v !== undefined ? Number(v) : undefined;
				break;
			case 'charset':
				f.charset = v !== undefined ? Number(v) : undefined;
				break;
			case 'scheme':
				f.scheme = v as FontSpec['scheme'];
				break;
			default:
				break;
		}
	}
	return f;
}

/** First child by local name (prefix-agnostic). */
export function childByLocal(el: XmlElement, local: string): XmlElement | undefined {
	return child(el, local) ?? elements(el).find((c) => localName(c.name) === local);
}

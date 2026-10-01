import type { RichTextRun } from '../types/cell';
import { isElement } from '../xml/dom';
import { decodeOoxmlEscapes, decodeXmlEntities } from '../xml/escape';
import { parseXml } from '../xml/parser';
import type { XmlElement } from '../xml/types';
import { parseFont } from './style-parts';
import { stripElementPrefix } from './xml-normalize';
import { kid, lname, ooxmlText } from './xml-utils';

/** A string item (`<si>`, `<is>`, comment `<text>`): plain text plus runs when formatted. */
export interface StringItem {
	text: string;
	runs?: RichTextRun[];
}

/**
 * Read a `CT_Rst` element: the plain `<t>` and the `<r>` runs, in order.
 * Phonetic runs (`<rPh>`) and `<phoneticPr>` are not part of the value and
 * are skipped. Runs are only reported when at least one carries formatting.
 */
export function parseRichElement(el: XmlElement): StringItem {
	let text = '';
	const runs: RichTextRun[] = [];
	let formatted = false;
	for (const c of el.children) {
		if (!isElement(c)) {
			continue;
		}
		const name = lname(c);
		if (name === 't') {
			const t = ooxmlText(c);
			text += t;
			runs.push({ text: t });
		} else if (name === 'r') {
			const t = ooxmlText(kid(c, 't'));
			const run: RichTextRun = { text: t };
			const rPr = kid(c, 'rPr');
			if (rPr) {
				const font = parseFont(rPr);
				if (Object.keys(font).length > 0) {
					run.font = font;
					formatted = true;
				}
			}
			text += t;
			runs.push(run);
		}
	}
	return formatted ? { text, runs } : { text };
}

function isTagNameEnd(c: number): boolean {
	return c === 62 || c === 47 || c === 32 || c === 9 || c === 10 || c === 13;
}

/** Decode raw element text: XML entities, then OOXML `_xHHHH_` escapes. */
export function decodeCellText(raw: string): string {
	if (raw.indexOf('<![CDATA[') !== -1) {
		// CDATA content carries no entities, but `_xHHHH_` escapes still apply.
		return raw
			.split(/<!\[CDATA\[([\s\S]*?)\]\]>/)
			.map((part, i) => (i % 2 === 1 ? decodeOoxmlEscapes(part) : decodeCellText(part)))
			.join('');
	}
	return decodeOoxmlEscapes(raw.indexOf('&') === -1 ? raw : decodeXmlEntities(raw));
}

/**
 * Read a string item straight from the document text. `elStart..elEnd` spans
 * the whole element and `contentStart..contentEnd` its content. Plain items
 * (a single `<t>`) are decoded without building a tree; anything else (runs,
 * phonetic runs, CDATA) goes through the XML parser.
 */
export function readStringItem(
	xml: string,
	elStart: number,
	contentStart: number,
	contentEnd: number,
	elEnd: number,
	prefix: string,
): StringItem {
	const p = prefix ? `${prefix}:` : '';
	const rTag = `<${p}r`;
	const r = xml.indexOf(rTag, contentStart);
	if (r === -1 || r >= contentEnd) {
		const tTag = `<${p}t`;
		let t = xml.indexOf(tTag, contentStart);
		while (t !== -1 && t < contentEnd && !isTagNameEnd(xml.charCodeAt(t + tTag.length))) {
			t = xml.indexOf(tTag, t + 1);
		}
		if (t === -1 || t >= contentEnd) {
			return { text: '' };
		}
		const gt = xml.indexOf('>', t);
		if (gt === -1 || gt >= contentEnd) {
			return { text: '' };
		}
		if (xml.charCodeAt(gt - 1) === 47) {
			return { text: '' };
		}
		const close = xml.indexOf(`</${p}t>`, gt);
		if (close !== -1 && close <= contentEnd) {
			const raw = xml.slice(gt + 1, close);
			if (raw.indexOf('<') === -1) {
				return { text: decodeCellText(raw) };
			}
		}
	}
	const el = parseXml(xml.slice(elStart, elEnd)).root;
	stripElementPrefix(el, prefix);
	return parseRichElement(el);
}

import { escapeXmlAttr, escapeXmlText } from './escape';
import type { XmlDocument, XmlElement, XmlNode } from './types';

/** The standard declaration Excel writes at the top of every part. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function writeNode(node: XmlNode, out: string[]): void {
	if (typeof node === 'string') {
		out.push(escapeXmlText(node));
		return;
	}
	if ('raw' in node) {
		out.push(node.raw);
		return;
	}
	writeElement(node, out);
}

function writeElement(el: XmlElement, out: string[]): void {
	let open = `<${el.name}`;
	for (const key in el.attrs) {
		const v = el.attrs[key];
		if (v === undefined) {
			continue;
		}
		open += ` ${key}="${escapeXmlAttr(v)}"`;
	}
	if (el.children.length === 0) {
		out.push(`${open}/>`);
		return;
	}
	out.push(`${open}>`);
	for (const child of el.children) {
		writeNode(child, out);
	}
	out.push(`</${el.name}>`);
}

/** Serialize a single element (no declaration). */
export function serializeElement(el: XmlElement): string {
	const out: string[] = [];
	writeElement(el, out);
	return out.join('');
}

/**
 * Serialize a document. When the document has no prolog the standard Excel
 * declaration is emitted.
 */
export function serializeXml(doc: XmlDocument | XmlElement): string {
	if ('root' in doc) {
		const prolog = doc.prolog.trim() ? doc.prolog : XML_DECLARATION;
		return prolog + serializeElement(doc.root);
	}
	return XML_DECLARATION + serializeElement(doc);
}

/**
 * Minimal ordered XML tree used for every OOXML part.
 *
 * Element and attribute names keep their namespace prefixes exactly as they
 * appear in the source (`x14ac:dyDescent`, `r:id`), so a parsed tree can be
 * serialized back without losing markup the engine does not understand.
 */

export interface XmlElement {
	/** Qualified tag name including any prefix, e.g. `c`, `x14:dataBar`. */
	name: string;
	/** Attributes in source order (object key order is insertion order). Values are decoded. */
	attrs: Record<string, string>;
	/** Child nodes in document order. Text nodes are plain (decoded) strings. */
	children: XmlNode[];
}

/** A comment, CDATA section or processing instruction kept verbatim for round-trip. */
export interface XmlRawNode {
	raw: string;
}

export type XmlNode = XmlElement | string | XmlRawNode;

/** A parsed document: optional prolog (XML declaration and PIs) plus the root element. */
export interface XmlDocument {
	/** Verbatim text that preceded the root element (XML declaration, processing instructions). */
	prolog: string;
	root: XmlElement;
}

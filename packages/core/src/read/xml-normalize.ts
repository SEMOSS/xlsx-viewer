import { isElement } from '../xml/dom';
import { decodeXmlEntities } from '../xml/escape';
import { NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import type { XmlDocument, XmlElement } from '../xml/types';

/**
 * Namespace normalisation for SpreadsheetML parts.
 *
 * Excel always writes the SpreadsheetML namespace as the default namespace,
 * but other producers bind it to a prefix (`<x:worksheet xmlns:x="...">`).
 * Readers work on unprefixed names, so the tree is rewritten: element names
 * lose the prefix bound to the main namespace, the root gains a matching
 * default-namespace declaration, and a non-standard prefix for the
 * relationships namespace is renamed to `r` (with an `xmlns:r` declaration).
 * The rewritten tree serializes to equivalent XML, so it is safe to keep as
 * round-trip source.
 */

const MAIN_NAMESPACES: ReadonlySet<string> = new Set([NS.main, NS.mainStrict]);
const REL_NAMESPACES: ReadonlySet<string> = new Set([NS.r, NS.rStrict]);

/** The root start tag of a document, located without building a tree. */
export interface RootTag {
	name: string;
	attrs: Record<string, string>;
	/** Offset of the `<` that opens the root tag. */
	start: number;
	/** Offset just past the `>` that closes the root start tag. */
	end: number;
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Parse the attributes of a start-tag body (the text after the element name). */
export function parseTagAttributes(source: string): Record<string, string> {
	const attrs: Record<string, string> = {};
	ATTR_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	while ((m = ATTR_RE.exec(source)) !== null) {
		attrs[m[1]] = decodeXmlEntities(m[2] ?? m[3] ?? '');
	}
	return attrs;
}

/** Offset just past the `>` closing the tag that starts at `lt`, honouring quoted attribute values. */
export function tagEnd(xml: string, lt: number): number {
	let quote = 0;
	for (let i = lt + 1; i < xml.length; i++) {
		const c = xml.charCodeAt(i);
		if (quote !== 0) {
			if (c === quote) {
				quote = 0;
			}
		} else if (c === 34 || c === 39) {
			quote = c;
		} else if (c === 62) {
			return i + 1;
		}
	}
	return -1;
}

/** Locate and parse the root element's start tag (skipping the prolog). */
export function scanRootTag(xml: string): RootTag | undefined {
	let pos = 0;
	for (;;) {
		const lt = xml.indexOf('<', pos);
		if (lt === -1) {
			return undefined;
		}
		const next = xml.charCodeAt(lt + 1);
		if (next === 63 /* ? */) {
			const e = xml.indexOf('?>', lt + 2);
			if (e === -1) {
				return undefined;
			}
			pos = e + 2;
			continue;
		}
		if (next === 33 /* ! */) {
			const e = xml.startsWith('<!--', lt)
				? xml.indexOf('-->', lt + 4) + 3
				: xml.indexOf('>', lt) + 1;
			if (e <= 0) {
				return undefined;
			}
			pos = e;
			continue;
		}
		const end = tagEnd(xml, lt);
		if (end === -1) {
			return undefined;
		}
		const body = xml.slice(lt + 1, xml.charCodeAt(end - 2) === 47 ? end - 2 : end - 1);
		const nameMatch = /^[^\s/>]+/.exec(body);
		const name = nameMatch ? nameMatch[0] : '';
		return { name, attrs: parseTagAttributes(body.slice(name.length)), start: lt, end };
	}
}

/** Prefix bound to the SpreadsheetML main namespace (`''` when it is the default namespace). */
export function spreadsheetPrefix(attrs: Record<string, string>): string {
	const def = attrs.xmlns;
	if (def !== undefined && MAIN_NAMESPACES.has(def)) {
		return '';
	}
	for (const key in attrs) {
		if (key.startsWith('xmlns:') && MAIN_NAMESPACES.has(attrs[key])) {
			return key.slice(6);
		}
	}
	return '';
}

/** The document uses ISO/IEC 29500 Strict namespaces. */
export function isStrictDocument(attrs: Record<string, string>): boolean {
	for (const key in attrs) {
		if (key === 'xmlns' || key.startsWith('xmlns:')) {
			if (attrs[key].startsWith('http://purl.oclc.org/ooxml/')) {
				return true;
			}
		}
	}
	return false;
}

function relationshipPrefix(attrs: Record<string, string>): string | undefined {
	for (const key in attrs) {
		if (key.startsWith('xmlns:') && REL_NAMESPACES.has(attrs[key])) {
			return key.slice(6);
		}
	}
	return undefined;
}

function renameTree(el: XmlElement, elementPrefix: string, attrPrefix: string): void {
	const stack: XmlElement[] = [el];
	while (stack.length > 0) {
		const cur = stack.pop()!;
		if (elementPrefix && cur.name.startsWith(elementPrefix)) {
			cur.name = cur.name.slice(elementPrefix.length);
		}
		if (attrPrefix) {
			let renamed: Record<string, string> | undefined;
			for (const key in cur.attrs) {
				if (key.startsWith(attrPrefix)) {
					renamed = {};
					break;
				}
			}
			if (renamed) {
				for (const key in cur.attrs) {
					const k = key.startsWith(attrPrefix) ? `r:${key.slice(attrPrefix.length)}` : key;
					renamed[k] = cur.attrs[key];
				}
				cur.attrs = renamed;
			}
		}
		for (const c of cur.children) {
			if (isElement(c)) {
				stack.push(c);
			}
		}
	}
}

/**
 * Rewrite a SpreadsheetML tree so the main namespace is the default
 * namespace and relationship attributes use the `r` prefix (see the module
 * comment). Returns the prefix that was stripped (`''` when none).
 */
export function normalizeSpreadsheetTree(root: XmlElement): string {
	const attrs = root.attrs;
	const prefix = spreadsheetPrefix(attrs);
	const relPrefix = relationshipPrefix(attrs);
	const renameRel = relPrefix !== undefined && relPrefix !== 'r' && attrs['xmlns:r'] === undefined;
	if (!prefix && !renameRel) {
		return '';
	}
	renameTree(root, prefix ? `${prefix}:` : '', renameRel ? `${relPrefix}:` : '');
	if (prefix && attrs.xmlns === undefined) {
		root.attrs = { xmlns: attrs[`xmlns:${prefix}`], ...root.attrs };
	}
	if (renameRel) {
		root.attrs['xmlns:r'] = root.attrs[`xmlns:${relPrefix}`];
	}
	return prefix;
}

/** Parse a SpreadsheetML part and normalise its namespaces. */
export function parseSpreadsheetXml(xml: string): XmlDocument {
	const doc = parseXml(xml);
	normalizeSpreadsheetTree(doc.root);
	return doc;
}

/** Strip a known element prefix from a parsed fragment (no namespace declarations involved). */
export function stripElementPrefix(el: XmlElement, prefix: string): void {
	if (prefix) {
		renameTree(el, `${prefix}:`, '');
	}
}

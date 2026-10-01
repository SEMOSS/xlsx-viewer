/**
 * Small XML helpers shared by the part writers: attribute merging, schema
 * ordered child replacement (aware of `mc:AlternateContent` and prefixed
 * roots), number formatting and text escaping in the forms Excel writes.
 */
import type { ColorSpec } from '../types/style';
import {
	elements,
	encodeOoxmlEscapes,
	escapeXmlAttr,
	escapeXmlText,
	isElement,
	localName,
	namePrefix,
	serializeElement,
	XML_DECLARATION,
} from '../xml';
import type { XmlElement, XmlNode } from '../xml/types';

export type AttrValue = string | number | boolean | undefined | null;
export type AttrMap = Record<string, AttrValue>;

/** Shortest exact decimal representation, with an upper-case exponent (`1.5E-7`). */
export function formatXmlNumber(n: number): string {
	if (n === 0) {
		return '0';
	}
	const s = String(n);
	const e = s.indexOf('e');
	return e === -1 ? s : `${s.slice(0, e)}E${s.slice(e + 1)}`;
}

/** Attribute value as text (`undefined` drops the attribute; booleans become `1` / `0`). */
export function attrText(v: AttrValue): string | undefined {
	if (v === undefined || v === null) {
		return undefined;
	}
	if (typeof v === 'boolean') {
		return v ? '1' : '0';
	}
	if (typeof v === 'number') {
		return formatXmlNumber(v);
	}
	return v;
}

/** `1` for true, nothing otherwise (for attributes whose default is false). */
export function flag(v: boolean | undefined): '1' | undefined {
	return v ? '1' : undefined;
}

/** `0` for false, nothing otherwise (for attributes whose default is true). */
export function unflag(v: boolean | undefined): '0' | undefined {
	return v === false ? '0' : undefined;
}

/**
 * Merge model-owned attributes into an original attribute set. Every key of
 * `values` is owned by the model: it is updated in place (keeping its
 * original position), removed when the value is `undefined`, or appended.
 * Attributes the model does not know are kept as they were.
 */
export function mergeAttrs(
	original: Record<string, string> | undefined,
	values: AttrMap,
): Record<string, string> {
	const out: Record<string, string> = {};
	if (original) {
		for (const k in original) {
			if (k in values) {
				const v = attrText(values[k]);
				if (v !== undefined) {
					out[k] = v;
				}
			} else {
				out[k] = original[k];
			}
		}
	}
	for (const k in values) {
		if (!(k in out)) {
			const v = attrText(values[k]);
			if (v !== undefined) {
				out[k] = v;
			}
		}
	}
	return out;
}

/** Serialize attributes (` k="v"` pairs) from an attribute map. */
export function attrString(attrs: AttrMap): string {
	let s = '';
	for (const k in attrs) {
		const v = attrText(attrs[k]);
		if (v !== undefined) {
			s += ` ${k}="${escapeXmlAttr(v)}"`;
		}
	}
	return s;
}

/** A self-closing or text element as a string. */
export function tag(name: string, attrs: AttrMap = {}, inner?: string): string {
	const a = attrString(attrs);
	return inner === undefined || inner === '' ? `<${name}${a}/>` : `<${name}${a}>${inner}</${name}>`;
}

// ── Text ──────────────────────────────────────────────────────────────────

function isSpace(c: number): boolean {
	return c === 32 || c === 9 || c === 10 || c === 13;
}

/** Text that needs `xml:space="preserve"` to keep its whitespace. */
export function needsPreserve(text: string): boolean {
	if (text.length === 0) {
		return false;
	}
	return (
		isSpace(text.charCodeAt(0)) ||
		isSpace(text.charCodeAt(text.length - 1)) ||
		text.indexOf('\n') !== -1 ||
		text.indexOf('\t') !== -1
	);
}

/** Escape cell / shared-string text: literal `_xHHHH_` sequences, XML specials, control chars. */
export function escapeCellText(text: string): string {
	return escapeXmlText(encodeOoxmlEscapes(text));
}

/** A `<t>` element (optionally prefixed) with whitespace preservation when needed. */
export function textElement(text: string, prefix = ''): string {
	const name = `${prefix}t`;
	const body = escapeCellText(text);
	return needsPreserve(text)
		? `<${name} xml:space="preserve">${body}</${name}>`
		: `<${name}>${body}</${name}>`;
}

// ── Colors ────────────────────────────────────────────────────────────────

/** Normalise an ARGB / RGB hex string to the 8-digit upper-case ARGB Excel expects. */
export function normalizeArgb(rgb: string): string {
	const hex = rgb.replace(/^#/, '').toUpperCase();
	if (/^[0-9A-F]{6}$/.test(hex)) {
		return `FF${hex}`;
	}
	return hex;
}

/** `CT_Color` attributes in the order Excel writes them. */
export function colorAttrs(c: ColorSpec): AttrMap {
	return {
		auto: c.auto ? true : undefined,
		indexed: c.indexed,
		rgb: c.rgb !== undefined ? normalizeArgb(c.rgb) : undefined,
		theme: c.theme,
		tint: c.tint !== undefined && c.tint !== 0 ? c.tint : undefined,
	};
}

/** A color element string such as `<color theme="1"/>`. */
export function colorTag(name: string, c: ColorSpec | undefined): string {
	return c ? tag(name, colorAttrs(c)) : '';
}

// ── Element order ─────────────────────────────────────────────────────────

/** Local name of a direct child that lives in the main namespace (unprefixed or `mainPrefix`). */
function mainLocal(node: XmlNode, mainPrefix: string): string | undefined {
	if (!isElement(node)) {
		return undefined;
	}
	const p = namePrefix(node.name);
	return p === '' || p === mainPrefix ? localName(node.name) : undefined;
}

function rankOf(node: XmlNode, order: readonly string[], mainPrefix: string): number {
	if (!isElement(node)) {
		return -1;
	}
	const local = mainLocal(node, mainPrefix);
	if (local !== undefined) {
		return order.indexOf(local);
	}
	if (localName(node.name) === 'AlternateContent') {
		for (const branch of elements(node)) {
			const first = elements(branch)[0];
			if (first) {
				return rankOf(first, order, mainPrefix);
			}
		}
	}
	return -1;
}

/** First direct child with a main-namespace local name. */
export function findChild(
	parent: XmlElement | undefined,
	name: string,
	mainPrefix = '',
): XmlElement | undefined {
	if (!parent) {
		return undefined;
	}
	for (const c of parent.children) {
		if (mainLocal(c, mainPrefix) === name) {
			return c as XmlElement;
		}
	}
	return undefined;
}

/** All direct children with a main-namespace local name. */
export function findChildren(
	parent: XmlElement | undefined,
	name: string,
	mainPrefix = '',
): XmlElement[] {
	const out: XmlElement[] = [];
	for (const c of parent?.children ?? []) {
		if (mainLocal(c, mainPrefix) === name) {
			out.push(c as XmlElement);
		}
	}
	return out;
}

/**
 * Replace every child named `name` with `replacement` (at the position of the
 * first one), or insert it at its schema position when absent. `null` (or an
 * empty array) removes the element. Unknown elements keep their positions;
 * `mc:AlternateContent` blocks rank as the element they wrap.
 */
export function setOrderedChild(
	parent: XmlElement,
	name: string,
	replacement: XmlElement | XmlElement[] | null,
	order: readonly string[],
	mainPrefix = '',
): void {
	const news = replacement === null ? [] : Array.isArray(replacement) ? replacement : [replacement];
	const kids = parent.children;
	if (kids.some((c) => mainLocal(c, mainPrefix) === name)) {
		const kept: XmlNode[] = [];
		let inserted = false;
		for (const c of kids) {
			if (mainLocal(c, mainPrefix) === name) {
				if (!inserted) {
					kept.push(...news);
					inserted = true;
				}
				continue;
			}
			kept.push(c);
		}
		parent.children = kept;
		return;
	}
	if (news.length === 0) {
		return;
	}
	const rank = order.indexOf(name);
	let at = kids.length;
	if (rank !== -1) {
		for (let i = 0; i < kids.length; i++) {
			if (rankOf(kids[i], order, mainPrefix) > rank) {
				at = i;
				break;
			}
		}
	}
	kids.splice(at, 0, ...news);
}

// ── Roots and namespaces ──────────────────────────────────────────────────

/**
 * Make `root` an unprefixed element in the default namespace `ns`. Returns the
 * prefix the source used for that namespace (so its children can still be
 * matched), or `''`.
 */
export function normalizeRoot(root: XmlElement, local: string, ns: string): string {
	const prefix = namePrefix(root.name);
	root.name = local;
	if (root.attrs.xmlns !== ns) {
		const rest = { ...root.attrs };
		delete rest.xmlns;
		root.attrs = { xmlns: ns, ...rest };
	}
	return prefix;
}

/** Prefix bound to a namespace URI on an element, if any. */
export function prefixForNamespace(el: XmlElement, uri: string): string | undefined {
	for (const k in el.attrs) {
		if (k.startsWith('xmlns:') && el.attrs[k] === uri) {
			return k.slice(6);
		}
	}
	return undefined;
}

/** Ensure a namespace is declared on `el`; returns the prefix to use for it. */
export function ensureNamespace(el: XmlElement, preferred: string, uri: string): string {
	const existing = prefixForNamespace(el, uri);
	if (existing) {
		return existing;
	}
	let name = preferred;
	for (let i = 1; el.attrs[`xmlns:${name}`] !== undefined; i++) {
		name = `${preferred}${i}`;
	}
	el.attrs[`xmlns:${name}`] = uri;
	return name;
}

function serializeNode(node: XmlNode): string {
	if (typeof node === 'string') {
		return escapeXmlText(node);
	}
	if (isElement(node)) {
		return serializeElement(node);
	}
	return node.raw;
}

/**
 * Serialize a document root, substituting pre-built XML strings for some of
 * its children (used to splice in the stream-built `sheetData`).
 */
export function serializeRootWith(
	root: XmlElement,
	substitutes: ReadonlyMap<XmlElement, string> = new Map(),
	declaration = true,
): string {
	const parts: string[] = [declaration ? XML_DECLARATION : '', `<${root.name}`];
	for (const k in root.attrs) {
		parts.push(` ${k}="${escapeXmlAttr(root.attrs[k])}"`);
	}
	if (root.children.length === 0) {
		parts.push('/>');
		return parts.join('');
	}
	parts.push('>');
	for (const c of root.children) {
		const sub = isElement(c) ? substitutes.get(c) : undefined;
		parts.push(sub ?? serializeNode(c));
	}
	parts.push(`</${root.name}>`);
	return parts.join('');
}

// ── Ids ───────────────────────────────────────────────────────────────────

/** A new upper-case GUID in braces, as Office writes ids (`{3F2504E0-...}`). */
export function newGuid(): string {
	const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
	if (c?.randomUUID) {
		return `{${c.randomUUID().toUpperCase()}}`;
	}
	const hex = '0123456789ABCDEF';
	let s = '';
	for (let i = 0; i < 32; i++) {
		s += hex[Math.floor(Math.random() * 16)];
	}
	return `{${s.slice(0, 8)}-${s.slice(8, 12)}-4${s.slice(13, 16)}-${s.slice(16, 20)}-${s.slice(20)}}`;
}

const GUID_RE = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;

/** The id is a braced GUID. */
export function isGuid(id: string | undefined): boolean {
	return id !== undefined && GUID_RE.test(id);
}

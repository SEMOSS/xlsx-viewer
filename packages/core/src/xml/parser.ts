import { decodeXmlEntities } from './escape';
import type { XmlDocument, XmlElement, XmlNode } from './types';

export class XmlParseError extends Error {
	constructor(
		message: string,
		readonly position: number,
	) {
		super(`${message} (at offset ${position})`);
		this.name = 'XmlParseError';
	}
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function parseAttributes(source: string, attrs: Record<string, string>): void {
	ATTR_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	while ((m = ATTR_RE.exec(source)) !== null) {
		const raw = m[2] !== undefined ? m[2] : m[3];
		attrs[m[1]] = decodeXmlEntities(raw);
	}
}

function isWhitespaceOnly(s: string): boolean {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c !== 32 && c !== 9 && c !== 10 && c !== 13) {
			return false;
		}
	}
	return true;
}

/**
 * Drop whitespace-only text nodes from elements that also contain element
 * children (indentation between tags). Whitespace that is an element's only
 * content, such as `<t xml:space="preserve"> </t>`, is kept.
 */
function trimFormattingWhitespace(el: XmlElement): void {
	const kids = el.children;
	if (kids.length < 2) {
		return;
	}
	let hasElement = false;
	for (const k of kids) {
		if (typeof k === 'object' && 'name' in k) {
			hasElement = true;
			break;
		}
	}
	if (!hasElement || el.attrs['xml:space'] === 'preserve') {
		return;
	}
	el.children = kids.filter((k) => typeof k !== 'string' || !isWhitespaceOnly(k));
}

/**
 * Parse an XML document into an ordered element tree.
 *
 * This is a small non-validating parser tailored to OOXML parts: it keeps
 * prefixed names verbatim, decodes entities, preserves comments and
 * processing instructions inside the tree as raw nodes, and treats CDATA as
 * text. It is iterative (no recursion), so deeply nested input is safe.
 */
export function parseXml(input: string): XmlDocument {
	let text = input;
	if (text.charCodeAt(0) === 0xfeff) {
		text = text.slice(1);
	}
	const len = text.length;
	let pos = 0;
	let prolog = '';
	let root: XmlElement | null = null;
	const stack: XmlElement[] = [];

	const pushNode = (node: XmlNode): void => {
		const parent = stack[stack.length - 1];
		if (parent) {
			parent.children.push(node);
		}
	};

	while (pos < len) {
		const lt = text.indexOf('<', pos);
		if (lt === -1) {
			// Trailing text after the root element is ignored; inside an element it is content.
			if (stack.length > 0) {
				pushNode(decodeXmlEntities(text.slice(pos)));
			}
			break;
		}
		if (lt > pos) {
			const chunk = text.slice(pos, lt);
			if (stack.length > 0) {
				pushNode(decodeXmlEntities(chunk));
			} else if (!root) {
				prolog += chunk;
			}
		}
		const next = text.charCodeAt(lt + 1);
		if (next === 63 /* ? */) {
			const end = text.indexOf('?>', lt + 2);
			if (end === -1) {
				throw new XmlParseError('Unterminated processing instruction', lt);
			}
			const raw = text.slice(lt, end + 2);
			if (stack.length > 0) {
				pushNode({ raw });
			} else if (!root) {
				prolog += raw;
			}
			pos = end + 2;
			continue;
		}
		if (next === 33 /* ! */) {
			if (text.startsWith('<!--', lt)) {
				const end = text.indexOf('-->', lt + 4);
				if (end === -1) {
					throw new XmlParseError('Unterminated comment', lt);
				}
				const raw = text.slice(lt, end + 3);
				if (stack.length > 0) {
					pushNode({ raw });
				} else if (!root) {
					prolog += raw;
				}
				pos = end + 3;
				continue;
			}
			if (text.startsWith('<![CDATA[', lt)) {
				const end = text.indexOf(']]>', lt + 9);
				if (end === -1) {
					throw new XmlParseError('Unterminated CDATA section', lt);
				}
				if (stack.length > 0) {
					pushNode(text.slice(lt + 9, end));
				}
				pos = end + 3;
				continue;
			}
			// DOCTYPE or other declaration: skip to the matching '>' (handles an internal subset).
			let depth = 0;
			let i = lt + 2;
			for (; i < len; i++) {
				const c = text.charCodeAt(i);
				if (c === 91 /* [ */) {
					depth++;
				} else if (c === 93 /* ] */) {
					depth--;
				} else if (c === 62 /* > */ && depth <= 0) {
					break;
				}
			}
			if (!root && stack.length === 0) {
				prolog += text.slice(lt, i + 1);
			}
			pos = i + 1;
			continue;
		}
		if (next === 47 /* / */) {
			const end = text.indexOf('>', lt + 2);
			if (end === -1) {
				throw new XmlParseError('Unterminated end tag', lt);
			}
			const name = text.slice(lt + 2, end).trim();
			const open = stack.pop();
			if (!open) {
				throw new XmlParseError(`Unexpected end tag </${name}>`, lt);
			}
			if (open.name !== name) {
				throw new XmlParseError(`Mismatched end tag </${name}>, expected </${open.name}>`, lt);
			}
			trimFormattingWhitespace(open);
			pos = end + 1;
			continue;
		}
		// Start tag: find the closing '>' while skipping quoted attribute values.
		let i = lt + 1;
		let quote = 0;
		for (; i < len; i++) {
			const c = text.charCodeAt(i);
			if (quote !== 0) {
				if (c === quote) {
					quote = 0;
				}
			} else if (c === 34 || c === 39) {
				quote = c;
			} else if (c === 62 /* > */) {
				break;
			}
		}
		if (i >= len) {
			throw new XmlParseError('Unterminated start tag', lt);
		}
		const selfClosing = text.charCodeAt(i - 1) === 47; /* / */
		const inner = text.slice(lt + 1, selfClosing ? i - 1 : i);
		let nameEnd = 0;
		while (nameEnd < inner.length) {
			const c = inner.charCodeAt(nameEnd);
			if (c === 32 || c === 9 || c === 10 || c === 13 || c === 47) {
				break;
			}
			nameEnd++;
		}
		const el: XmlElement = { name: inner.slice(0, nameEnd), attrs: {}, children: [] };
		if (nameEnd < inner.length) {
			parseAttributes(inner.slice(nameEnd), el.attrs);
		}
		if (stack.length === 0) {
			if (root) {
				throw new XmlParseError('Multiple root elements', lt);
			}
			root = el;
		} else {
			pushNode(el);
		}
		if (!selfClosing) {
			stack.push(el);
		}
		pos = i + 1;
	}
	if (stack.length > 0) {
		throw new XmlParseError(`Unclosed element <${stack[stack.length - 1].name}>`, len);
	}
	if (!root) {
		throw new XmlParseError('No root element', 0);
	}
	return { prolog, root };
}

/** Parse and return only the root element. */
export function parseXmlRoot(input: string): XmlElement {
	return parseXml(input).root;
}

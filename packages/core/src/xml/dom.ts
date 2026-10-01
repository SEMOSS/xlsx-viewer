import type { XmlElement, XmlNode } from './types';

/** Type guard: node is an element (not text or raw). */
export function isElement(node: XmlNode | undefined): node is XmlElement {
	return typeof node === 'object' && node !== null && 'name' in node;
}

/** Local part of a qualified name (`x14:dataBar` -> `dataBar`). */
export function localName(name: string): string {
	const i = name.indexOf(':');
	return i === -1 ? name : name.slice(i + 1);
}

/** Prefix of a qualified name (`x14:dataBar` -> `x14`, `c` -> `''`). */
export function namePrefix(name: string): string {
	const i = name.indexOf(':');
	return i === -1 ? '' : name.slice(0, i);
}

/** Element children only. */
export function elements(el: XmlElement | undefined): XmlElement[] {
	if (!el) {
		return [];
	}
	const out: XmlElement[] = [];
	for (const c of el.children) {
		if (isElement(c)) {
			out.push(c);
		}
	}
	return out;
}

/** First child element with the exact qualified name. */
export function child(el: XmlElement | undefined, name: string): XmlElement | undefined {
	if (!el) {
		return undefined;
	}
	for (const c of el.children) {
		if (isElement(c) && c.name === name) {
			return c;
		}
	}
	return undefined;
}

/** All child elements with the exact qualified name. */
export function childrenNamed(el: XmlElement | undefined, name: string): XmlElement[] {
	if (!el) {
		return [];
	}
	const out: XmlElement[] = [];
	for (const c of el.children) {
		if (isElement(c) && c.name === name) {
			out.push(c);
		}
	}
	return out;
}

/** First child element whose local name matches (ignores the prefix). */
export function childLocal(el: XmlElement | undefined, local: string): XmlElement | undefined {
	if (!el) {
		return undefined;
	}
	for (const c of el.children) {
		if (isElement(c) && localName(c.name) === local) {
			return c;
		}
	}
	return undefined;
}

/** All child elements whose local name matches. */
export function childrenLocal(el: XmlElement | undefined, local: string): XmlElement[] {
	if (!el) {
		return [];
	}
	const out: XmlElement[] = [];
	for (const c of el.children) {
		if (isElement(c) && localName(c.name) === local) {
			out.push(c);
		}
	}
	return out;
}

/** Follow a path of exact child names: `path(root, 'sheetViews', 'sheetView')`. */
export function path(el: XmlElement | undefined, ...names: string[]): XmlElement | undefined {
	let cur = el;
	for (const n of names) {
		cur = child(cur, n);
		if (!cur) {
			return undefined;
		}
	}
	return cur;
}

/** Depth-first search for the first descendant with the given local name. */
export function findDescendantLocal(el: XmlElement, local: string): XmlElement | undefined {
	const stack: XmlElement[] = [el];
	while (stack.length) {
		const cur = stack.pop()!;
		for (let i = cur.children.length - 1; i >= 0; i--) {
			const c = cur.children[i];
			if (isElement(c)) {
				if (localName(c.name) === local) {
					return c;
				}
				stack.push(c);
			}
		}
	}
	return undefined;
}

/** All descendants with the given local name, in document order. */
export function findAllDescendantsLocal(el: XmlElement, local: string): XmlElement[] {
	const out: XmlElement[] = [];
	const visit = (node: XmlElement): void => {
		for (const c of node.children) {
			if (isElement(c)) {
				if (localName(c.name) === local) {
					out.push(c);
				}
				visit(c);
			}
		}
	};
	visit(el);
	return out;
}

/** Concatenated text of all descendant text nodes. */
export function textContent(el: XmlElement | undefined): string {
	if (!el) {
		return '';
	}
	if (el.children.length === 1 && typeof el.children[0] === 'string') {
		return el.children[0];
	}
	let out = '';
	for (const c of el.children) {
		if (typeof c === 'string') {
			out += c;
		} else if (isElement(c)) {
			out += textContent(c);
		}
	}
	return out;
}

export function attr(el: XmlElement | undefined, name: string): string | undefined {
	return el?.attrs[name];
}

/** Numeric attribute, or `fallback` when missing or not a number. */
export function attrNum(el: XmlElement | undefined, name: string): number | undefined;
export function attrNum(el: XmlElement | undefined, name: string, fallback: number): number;
export function attrNum(
	el: XmlElement | undefined,
	name: string,
	fallback?: number,
): number | undefined {
	const v = el?.attrs[name];
	if (v === undefined || v === '') {
		return fallback;
	}
	const n = Number(v);
	return Number.isFinite(n) ? n : fallback;
}

/** Boolean attribute (`1`, `true`, `on` are true; `0`, `false`, `off` are false). */
export function attrBool(el: XmlElement | undefined, name: string): boolean | undefined;
export function attrBool(el: XmlElement | undefined, name: string, fallback: boolean): boolean;
export function attrBool(
	el: XmlElement | undefined,
	name: string,
	fallback?: boolean,
): boolean | undefined {
	const v = el?.attrs[name];
	if (v === undefined) {
		return fallback;
	}
	const s = v.toLowerCase();
	if (s === '1' || s === 'true' || s === 'on') {
		return true;
	}
	if (s === '0' || s === 'false' || s === 'off') {
		return false;
	}
	return fallback;
}

/** `val` attribute of a child element, e.g. `<b val="0"/>`. */
export function childVal(el: XmlElement | undefined, name: string): string | undefined {
	return child(el, name)?.attrs.val;
}

/**
 * Boolean "toggle" child such as `<b/>` or `<b val="0"/>`: present without
 * `val` means true.
 */
export function childFlag(el: XmlElement | undefined, name: string): boolean | undefined {
	const c = child(el, name);
	if (!c) {
		return undefined;
	}
	const v = c.attrs.val;
	if (v === undefined) {
		return true;
	}
	return v === '1' || v.toLowerCase() === 'true';
}

// ── Building ───────────────────────────────────────────────────────────────

type AttrInput = Record<string, string | number | boolean | undefined | null>;

/**
 * Create an element. Attribute values are stringified; `undefined`/`null`
 * are skipped and booleans become `1` / `0`.
 */
export function el(
	name: string,
	attrs: AttrInput = {},
	children: Array<XmlNode | undefined | null | false> = [],
): XmlElement {
	const out: Record<string, string> = {};
	for (const k in attrs) {
		const v = attrs[k];
		if (v === undefined || v === null) {
			continue;
		}
		out[k] = typeof v === 'boolean' ? (v ? '1' : '0') : String(v);
	}
	const kids: XmlNode[] = [];
	for (const c of children) {
		if (c !== undefined && c !== null && c !== false) {
			kids.push(c);
		}
	}
	return { name, attrs: out, children: kids };
}

/** Element containing a single text node. */
export function textEl(name: string, text: string, attrs: AttrInput = {}): XmlElement {
	return el(name, attrs, [text]);
}

/** Deep clone an element tree. */
export function cloneElement(src: XmlElement): XmlElement {
	return {
		name: src.name,
		attrs: { ...src.attrs },
		children: src.children.map((c) =>
			typeof c === 'string' ? c : isElement(c) ? cloneElement(c) : { raw: c.raw },
		),
	};
}

/** Remove all children with the given exact name; returns the removed elements. */
export function removeChildren(parent: XmlElement, name: string): XmlElement[] {
	const removed: XmlElement[] = [];
	parent.children = parent.children.filter((c) => {
		if (isElement(c) && c.name === name) {
			removed.push(c);
			return false;
		}
		return true;
	});
	return removed;
}

/**
 * Replace (or insert) the child named `name` keeping schema order: `order` is
 * the list of allowed child names in sequence. Passing `null` removes it.
 * Every existing child named `name` is replaced by the single new element (or
 * `replacements` when several elements of the same name are needed).
 */
export function setChildInOrder(
	parent: XmlElement,
	name: string,
	replacement: XmlElement | XmlElement[] | null,
	order: readonly string[],
): void {
	const news = replacement === null ? [] : Array.isArray(replacement) ? replacement : [replacement];
	const firstExisting = parent.children.findIndex((c) => isElement(c) && c.name === name);
	if (firstExisting !== -1) {
		const kept: XmlNode[] = [];
		let inserted = false;
		for (let i = 0; i < parent.children.length; i++) {
			const c = parent.children[i];
			if (isElement(c) && c.name === name) {
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
	let insertAt = parent.children.length;
	if (rank !== -1) {
		for (let i = 0; i < parent.children.length; i++) {
			const c = parent.children[i];
			if (!isElement(c)) {
				continue;
			}
			const r = order.indexOf(c.name);
			if (r > rank) {
				insertAt = i;
				break;
			}
		}
	}
	parent.children.splice(insertAt, 0, ...news);
}

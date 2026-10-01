import { isElement, localName, textContent } from '../xml/dom';
import { decodeOoxmlEscapes } from '../xml/escape';
import type { XmlElement } from '../xml/types';

/** Small, allocation-light DOM helpers shared by the part readers. */

/** First child whose local name matches (prefix ignored). */
export function kid(el: XmlElement | undefined, local: string): XmlElement | undefined {
	if (!el) {
		return undefined;
	}
	for (const c of el.children) {
		if (isElement(c) && (c.name === local || localName(c.name) === local)) {
			return c;
		}
	}
	return undefined;
}

/** All children whose local name matches (prefix ignored). */
export function kids(el: XmlElement | undefined, local: string): XmlElement[] {
	const out: XmlElement[] = [];
	if (!el) {
		return out;
	}
	for (const c of el.children) {
		if (isElement(c) && (c.name === local || localName(c.name) === local)) {
			out.push(c);
		}
	}
	return out;
}

/** First element child, whatever its name. */
export function firstElement(el: XmlElement | undefined): XmlElement | undefined {
	if (!el) {
		return undefined;
	}
	for (const c of el.children) {
		if (isElement(c)) {
			return c;
		}
	}
	return undefined;
}

/** Local name of an element. */
export function lname(el: XmlElement): string {
	return localName(el.name);
}

/** Numeric attribute (`undefined` when absent or not a finite number). */
export function num(el: XmlElement | undefined, name: string): number | undefined {
	const v = el?.attrs[name];
	if (v === undefined || v === '') {
		return undefined;
	}
	const n = Number(v);
	return Number.isFinite(n) ? n : undefined;
}

/** Integer attribute (`undefined` when absent or invalid). */
export function int(el: XmlElement | undefined, name: string): number | undefined {
	const n = num(el, name);
	return n === undefined ? undefined : Math.trunc(n);
}

/** Boolean attribute: `1`/`true`/`on` and `0`/`false`/`off` (`undefined` when absent or invalid). */
export function bool(el: XmlElement | undefined, name: string): boolean | undefined {
	const v = el?.attrs[name];
	if (v === undefined) {
		return undefined;
	}
	return parseBoolText(v);
}

export function parseBoolText(v: string): boolean | undefined {
	const s = v.trim().toLowerCase();
	if (s === '1' || s === 'true' || s === 'on') {
		return true;
	}
	if (s === '0' || s === 'false' || s === 'off') {
		return false;
	}
	return undefined;
}

/** String attribute (`undefined` when absent). */
export function str(el: XmlElement | undefined, name: string): string | undefined {
	return el?.attrs[name];
}

/**
 * Toggle element such as `<b/>` or `<b val="0"/>` (matched by local name):
 * present without `val` means true; absent means `undefined`.
 */
export function flag(el: XmlElement | undefined, local: string): boolean | undefined {
	const c = kid(el, local);
	if (!c) {
		return undefined;
	}
	const v = c.attrs.val;
	return v === undefined ? true : (parseBoolText(v) ?? true);
}

/** `val` attribute of a child (matched by local name). */
export function childVal(el: XmlElement | undefined, local: string): string | undefined {
	return kid(el, local)?.attrs.val;
}

/** Numeric `val` attribute of a child. */
export function childNum(el: XmlElement | undefined, local: string): number | undefined {
	return num(kid(el, local), 'val');
}

/** Text content with the OOXML `_xHHHH_` escapes decoded. */
export function ooxmlText(el: XmlElement | undefined): string {
	return decodeOoxmlEscapes(textContent(el));
}

/**
 * Relationship attribute such as `r:id` / `r:embed`. Falls back to any
 * prefixed attribute with the same local name, for parts that bind the
 * relationships namespace to another prefix on an inner element.
 */
export function relAttr(el: XmlElement | undefined, local = 'id'): string | undefined {
	if (!el) {
		return undefined;
	}
	const direct = el.attrs[`r:${local}`];
	if (direct !== undefined) {
		return direct;
	}
	const suffix = `:${local}`;
	for (const key in el.attrs) {
		if (key.endsWith(suffix) && !key.startsWith('xmlns')) {
			return el.attrs[key];
		}
	}
	return undefined;
}

/** Remove `undefined` properties (keeps model objects tidy for deep equality and dedupe keys). */
export function compact<T extends object>(obj: T): T {
	for (const key of Object.keys(obj) as Array<keyof T>) {
		if (obj[key] === undefined) {
			delete obj[key];
		}
	}
	return obj;
}

/** Pick a value from a fixed set of literals, or `undefined`. */
export function oneOf<T extends string>(
	value: string | undefined,
	allowed: readonly T[],
): T | undefined {
	return value !== undefined && (allowed as readonly string[]).includes(value)
		? (value as T)
		: undefined;
}

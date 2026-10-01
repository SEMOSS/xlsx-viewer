/**
 * Checks for re-emitting original markup outside its original part: the
 * subtree must declare every prefix it uses (or find it in the target
 * root's scope) and must not reference relationships of its old part.
 */
import { elements, namePrefix } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';

/** Namespace prefixes declared on an element (`prefix -> uri`). */
export function prefixScope(attrs: Record<string, string>): Map<string, string> {
	const out = new Map<string, string>();
	for (const k in attrs) {
		if (k.startsWith('xmlns:')) {
			out.set(k.slice(6), attrs[k]);
		}
	}
	return out;
}

const REL_NAMESPACES = new Set<string>([NS.r, NS.rStrict]);

/**
 * The markup can be placed under a root declaring `scope`: all its prefixes
 * resolve and it carries no relationship-id attributes (`r:embed`, `r:id`, ...).
 */
export function isSelfContained(el: XmlElement, scope: ReadonlyMap<string, string>): boolean {
	let local = scope;
	for (const k in el.attrs) {
		if (k.startsWith('xmlns:')) {
			if (local === scope) {
				local = new Map(scope);
			}
			(local as Map<string, string>).set(k.slice(6), el.attrs[k]);
		}
	}
	const resolves = (qname: string): boolean => {
		const p = namePrefix(qname);
		return p === '' || p === 'xml' || local.has(p);
	};
	if (!resolves(el.name)) {
		return false;
	}
	for (const k in el.attrs) {
		if (k === 'xmlns' || k.startsWith('xmlns:')) {
			continue;
		}
		const p = namePrefix(k);
		if (!resolves(k) || (p && REL_NAMESPACES.has(local.get(p) ?? ''))) {
			return false;
		}
	}
	return elements(el).every((c) => isSelfContained(c, local));
}

/** Merging model-owned `<ext>` entries into a worksheet `<extLst>`, keeping the others. */
import { cloneElement, el, isElement, localName } from '../xml';
import type { XmlElement, XmlNode } from '../xml/types';
import { CF_EXT_URI } from './conditional-formats';
import { DV_EXT_URI } from './data-validations';
import { SPARKLINE_EXT_URI } from './sparklines';

/** Extensions the writer owns, in the order Excel writes them. */
export const MANAGED_EXT_URIS = [CF_EXT_URI, DV_EXT_URI, SPARKLINE_EXT_URI].map((u) =>
	u.toLowerCase(),
);

function extUri(node: XmlNode): string | undefined {
	return isElement(node) && localName(node.name) === 'ext'
		? node.attrs.uri?.toLowerCase()
		: undefined;
}

/**
 * Replace (or insert, or remove when `null`) the managed extensions. New ones
 * go before the first extension that Excel writes after them.
 */
export function mergeExtLst(
	existing: XmlElement | undefined,
	managed: ReadonlyArray<readonly [string, XmlElement | null]>,
): XmlElement | null {
	const out = existing ? cloneElement(existing) : el('extLst');
	for (const [uri, ext] of managed) {
		const key = uri.toLowerCase();
		const idx = out.children.findIndex((c) => extUri(c) === key);
		if (idx >= 0) {
			if (ext) {
				out.children[idx] = ext;
			} else {
				out.children.splice(idx, 1);
			}
			continue;
		}
		if (!ext) {
			continue;
		}
		const rank = MANAGED_EXT_URIS.indexOf(key);
		let at = out.children.length;
		for (let i = 0; i < out.children.length; i++) {
			const u = extUri(out.children[i]);
			if (u === undefined) {
				continue;
			}
			const r = MANAGED_EXT_URIS.indexOf(u);
			if (r === -1 || r > rank) {
				at = i;
				break;
			}
		}
		out.children.splice(at, 0, ext);
	}
	return out.children.some(isElement) ? out : null;
}

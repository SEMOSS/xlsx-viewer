import { findDescendantLocal } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseSpreadsheetXml } from './xml-normalize';
import { bool, int, kid, kids, str } from './xml-utils';

/** Metadata type Excel uses for dynamic-array formulas. */
const DYNAMIC_ARRAY_TYPE = 'XLDAPR';

/**
 * `xl/metadata.xml` -> the set of cell-metadata indices (the 1-based `cm`
 * attribute of `<c>`) that mark dynamic-array formulas. A cell block points
 * at a metadata type (`rc/@t`, 1-based) and at an entry of that type's
 * `futureMetadata` list (`rc/@v`, 0-based) holding `dynamicArrayProperties`.
 */
export function parseDynamicArrayMetadata(xml: string): Set<number> {
	const root = parseSpreadsheetXml(xml).root;
	const typeNames = kids(kid(root, 'metadataTypes'), 'metadataType').map(
		(t) => str(t, 'name') ?? '',
	);
	const future = new Map<string, XmlElement[]>();
	for (const fm of kids(root, 'futureMetadata')) {
		future.set(str(fm, 'name') ?? '', kids(fm, 'bk'));
	}
	const out = new Set<number>();
	kids(kid(root, 'cellMetadata'), 'bk').forEach((bk, i) => {
		for (const rc of kids(bk, 'rc')) {
			const typeName = typeNames[(int(rc, 't') ?? 0) - 1];
			if (typeName !== DYNAMIC_ARRAY_TYPE) {
				continue;
			}
			const entry = future.get(typeName)?.[int(rc, 'v') ?? 0];
			const props = entry ? findDescendantLocal(entry, 'dynamicArrayProperties') : undefined;
			if (!props || bool(props, 'fDynamic') !== false) {
				out.add(i + 1);
			}
		}
	});
	return out;
}

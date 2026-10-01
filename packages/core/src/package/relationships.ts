import type { Relationship } from '../types/sheet';
import { childrenLocal, el, parseXmlRoot, serializeXml } from '../xml';
import { NS } from '../xml/namespaces';

/** Parse a `.rels` part. */
export function parseRelationships(xml: string): Relationship[] {
	const root = parseXmlRoot(xml);
	return childrenLocal(root, 'Relationship').map((r) => ({
		id: r.attrs.Id ?? '',
		type: r.attrs.Type ?? '',
		target: r.attrs.Target ?? '',
		targetMode: r.attrs.TargetMode === 'External' ? 'External' : undefined,
	}));
}

/** Serialize relationships to a `.rels` part. */
export function serializeRelationships(rels: readonly Relationship[]): string {
	return serializeXml(
		el(
			'Relationships',
			{ xmlns: NS.pkgRels },
			rels.map((r) =>
				el('Relationship', {
					Id: r.id,
					Type: r.type,
					Target: r.target,
					TargetMode: r.targetMode === 'External' ? 'External' : undefined,
				}),
			),
		),
	);
}

/** Next free `rIdN` id. */
export function nextRelId(rels: readonly Relationship[]): string {
	let max = 0;
	for (const r of rels) {
		const m = /^rId(\d+)$/.exec(r.id);
		if (m) {
			max = Math.max(max, Number(m[1]));
		}
	}
	return `rId${max + 1}`;
}

/** First relationship of a type (exact match on the type URI suffix after the last slash). */
export function findRelByType(
	rels: readonly Relationship[],
	type: string,
): Relationship | undefined {
	const suffix = type.slice(type.lastIndexOf('/') + 1);
	return (
		rels.find((r) => r.type === type) ??
		rels.find((r) => r.type.slice(r.type.lastIndexOf('/') + 1) === suffix)
	);
}

/** All relationships of a type (matching on the type suffix, which also covers Strict URIs). */
export function findRelsByType(rels: readonly Relationship[], type: string): Relationship[] {
	const suffix = type.slice(type.lastIndexOf('/') + 1);
	return rels.filter(
		(r) => r.type === type || r.type.slice(r.type.lastIndexOf('/') + 1) === suffix,
	);
}

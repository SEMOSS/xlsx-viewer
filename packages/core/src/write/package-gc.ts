/**
 * Removal of parts orphaned by the save: a part that was reachable through
 * relationships in the source package but no longer is in the output
 * (deleted sheets with their drawings, comments, tables, charts and images,
 * replaced parts, a dropped calc chain) is removed with its content type.
 * Parts that were already unreachable in the source are left alone.
 */
import { resolveRelTarget } from '../package';
import type { Relationship } from '../types/sheet';
import type { WriteContext } from './context';

async function reachable(
	getRels: (part: string) => Promise<Relationship[]>,
	exists: (part: string) => boolean,
): Promise<Map<string, string>> {
	const seen = new Map<string, string>();
	const queue: string[] = [''];
	while (queue.length) {
		const part = queue.pop()!;
		let rels: Relationship[];
		try {
			rels = await getRels(part);
		} catch {
			continue;
		}
		for (const r of rels) {
			if (r.targetMode === 'External' || !r.target || r.target.startsWith('#')) {
				continue;
			}
			const target = resolveRelTarget(part, r.target.split('#')[0]);
			const key = target.toLowerCase();
			if (seen.has(key) || !exists(target)) {
				continue;
			}
			seen.set(key, target);
			queue.push(target);
		}
	}
	return seen;
}

/** Remove parts that the save disconnected; returns the removed paths. */
export async function removeOrphans(ctx: WriteContext): Promise<string[]> {
	const source = ctx.source;
	if (!source) {
		return [];
	}
	const before = await reachable(
		(p) => source.rels(p),
		(p) => source.has(p),
	);
	const after = await reachable(
		(p) => ctx.rels(p),
		(p) => ctx.has(p),
	);
	const removed: string[] = [];
	for (const [key, path] of before) {
		if (!after.has(key) && ctx.has(path)) {
			ctx.remove(path);
			removed.push(path);
		}
	}
	return removed;
}

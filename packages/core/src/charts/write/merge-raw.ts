/**
 * Re-insert markup the model does not cover (extension lists, external data,
 * 3-D view settings, error bars, ...) from the original chart XML into a
 * regenerated element, keeping schema order. This keeps edited charts as
 * faithful as possible without modelling every corner of the chart schema.
 */
import { cloneElement, isElement, localName } from '../../xml/dom';
import type { XmlElement } from '../../xml/types';
import { CHART_ORDER } from './schema';

export interface MergeSpec {
	/** Key into {@link CHART_ORDER}. */
	order: string;
	/** Local names the writer owns: raw children with these names are never copied. */
	owned: readonly string[];
	/** Owned names whose original element replaces the writer's default when present. */
	preferRaw?: readonly string[];
}

/**
 * Copy children of `raw` that the writer does not own into `generated`, at
 * their schema position. Returns `generated` (mutated) for chaining.
 */
export function mergeRawChildren(
	generated: XmlElement,
	raw: XmlElement | undefined,
	spec: MergeSpec,
): XmlElement {
	if (!raw) {
		return generated;
	}
	const order = CHART_ORDER[spec.order] ?? [];
	const rank = (local: string): number => order.indexOf(local);
	const owned = new Set(spec.owned);
	const preferRaw = new Set(spec.preferRaw ?? []);
	const generatedNames = new Set(
		generated.children.filter(isElement).map((c) => localName(c.name)),
	);
	let lastRank = -1;
	for (const c of raw.children) {
		if (!isElement(c)) {
			continue;
		}
		const local = localName(c.name);
		const r = rank(local);
		const effective = r >= 0 ? r : lastRank + 0.5;
		if (r >= 0) {
			lastRank = r;
		}
		if (preferRaw.has(local)) {
			const i = generated.children.findIndex((g) => isElement(g) && localName(g.name) === local);
			if (i >= 0) {
				generated.children[i] = cloneElement(c);
				continue;
			}
			if (owned.has(local)) {
				continue;
			}
		} else if (owned.has(local) || generatedNames.has(local)) {
			continue;
		}
		insertByRank(generated, cloneElement(c), effective, rank);
	}
	return generated;
}

function insertByRank(
	parent: XmlElement,
	node: XmlElement,
	nodeRank: number,
	rank: (local: string) => number,
): void {
	let at = parent.children.length;
	for (let i = 0; i < parent.children.length; i++) {
		const c = parent.children[i];
		if (isElement(c) && rank(localName(c.name)) > nodeRank) {
			at = i;
			break;
		}
	}
	parent.children.splice(at, 0, node);
}

/** Child elements of `raw` with a local name, for pairing regenerated elements with originals. */
export function rawChildren(raw: XmlElement | undefined, local: string): XmlElement[] {
	if (!raw) {
		return [];
	}
	return raw.children.filter((c): c is XmlElement => isElement(c) && localName(c.name) === local);
}

/** First child of `raw` with a local name. */
export function rawChild(raw: XmlElement | undefined, local: string): XmlElement | undefined {
	return rawChildren(raw, local)[0];
}

/** Child of `raw` whose `c:idx` / `c:axId` (given by `key`) equals `value`. */
export function rawChildByKey(
	raw: XmlElement | undefined,
	local: string,
	key: string,
	value: number,
): XmlElement | undefined {
	for (const c of rawChildren(raw, local)) {
		const k = rawChild(c, key);
		if (k && Number(k.attrs.val) === value) {
			return c;
		}
	}
	return undefined;
}

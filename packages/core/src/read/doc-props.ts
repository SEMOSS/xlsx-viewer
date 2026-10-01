import { findRelByType } from '../package/relationships';
import type { DocumentProperties } from '../types/workbook';
import { isElement, textContent } from '../xml/dom';
import { REL } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import type { XmlElement } from '../xml/types';
import { attempt, type ReadContext } from './context';
import { readPartText } from './part-io';
import { firstElement, kid, kids, lname, parseBoolText, str } from './xml-utils';

/** Document properties: `docProps/core.xml`, `app.xml` and `custom.xml`. */

const CORE_FIELDS: ReadonlyArray<[string, keyof DocumentProperties]> = [
	['title', 'title'],
	['subject', 'subject'],
	['creator', 'creator'],
	['keywords', 'keywords'],
	['description', 'description'],
	['lastModifiedBy', 'lastModifiedBy'],
	['category', 'category'],
	['created', 'created'],
	['modified', 'modified'],
];

function setText(
	props: DocumentProperties,
	key: keyof DocumentProperties,
	el: XmlElement | undefined,
): void {
	if (!el) {
		return;
	}
	const text = textContent(el).trim();
	if (text) {
		(props as Record<string, unknown>)[key] = text;
	}
}

/** `cp:coreProperties` (Dublin Core title, creator, dates, ...). Matched by local name. */
export function parseCoreProperties(xml: string, props: DocumentProperties): void {
	const root = parseXml(xml).root;
	const byLocal = new Map<string, XmlElement>();
	for (const c of root.children) {
		if (isElement(c) && !byLocal.has(lname(c))) {
			byLocal.set(lname(c), c);
		}
	}
	for (const [local, key] of CORE_FIELDS) {
		setText(props, key, byLocal.get(local));
	}
}

/** Extended properties (`Application`, `AppVersion`, `Company`, `Manager`). */
export function parseAppProperties(xml: string, props: DocumentProperties): void {
	const root = parseXml(xml).root;
	setText(props, 'application', kid(root, 'Application'));
	setText(props, 'appVersion', kid(root, 'AppVersion'));
	setText(props, 'company', kid(root, 'Company'));
	setText(props, 'manager', kid(root, 'Manager'));
}

const NUMERIC_TYPES: ReadonlySet<string> = new Set([
	'i1',
	'i2',
	'i4',
	'i8',
	'int',
	'ui1',
	'ui2',
	'ui4',
	'ui8',
	'uint',
	'r4',
	'r8',
	'decimal',
]);

/** Custom properties: each `<property name>` holds one typed `vt:*` value. */
export function parseCustomProperties(xml: string, props: DocumentProperties): void {
	const root = parseXml(xml).root;
	const custom: NonNullable<DocumentProperties['custom']> = [];
	for (const p of kids(root, 'property')) {
		const name = str(p, 'name');
		const valueEl = firstElement(p);
		if (!name || !valueEl) {
			continue;
		}
		const type = lname(valueEl);
		const text = textContent(valueEl);
		let value: string | number | boolean = text;
		if (NUMERIC_TYPES.has(type)) {
			const n = Number(text);
			value = Number.isFinite(n) ? n : text;
		} else if (type === 'bool') {
			value = parseBoolText(text) ?? false;
		}
		custom.push({ name, value, type });
	}
	if (custom.length > 0) {
		props.custom = custom;
	}
}

const SOURCES: ReadonlyArray<[string, string, (xml: string, props: DocumentProperties) => void]> = [
	[REL.coreProperties, 'docProps/core.xml', parseCoreProperties],
	[REL.extendedProperties, 'docProps/app.xml', parseAppProperties],
	[REL.customProperties, 'docProps/custom.xml', parseCustomProperties],
];

/** Read the document properties parts found through the package relationships. */
export async function readDocumentProperties(ctx: ReadContext): Promise<DocumentProperties> {
	const props: DocumentProperties = {};
	const rels = await ctx.pkg.rels('').catch(() => []);
	for (const [type, fallback, parse] of SOURCES) {
		const rel = findRelByType(rels, type);
		const path = rel && rel.targetMode !== 'External' ? ctx.pkg.resolve('', rel) : fallback;
		if (!ctx.pkg.has(path)) {
			continue;
		}
		await attempt(
			ctx,
			path,
			'document properties',
			async () => {
				const xml = await readPartText(ctx.pkg, path);
				if (xml !== undefined) {
					parse(xml, props);
				}
			},
			undefined,
		);
	}
	return props;
}

/**
 * Document properties: `docProps/core.xml` (modified time updated on every
 * save that changes the package), `docProps/app.xml` (sheet titles) and
 * `docProps/custom.xml`.
 */
import type { DocumentProperties } from '../types/workbook';
import {
	cloneElement,
	el,
	elements,
	localName,
	parseXml,
	serializeXml,
	textContent,
	textEl,
} from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';

/** W3CDTF timestamp without fractional seconds (`2024-01-31T12:00:00Z`). */
export function w3cdtf(value: Date | string): string {
	const d = typeof value === 'string' ? new Date(value) : value;
	if (Number.isNaN(d.getTime())) {
		return typeof value === 'string' ? value : new Date(0).toISOString().replace(/\.\d{3}Z$/, 'Z');
	}
	return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

const CORE_FIELDS: Array<[keyof DocumentProperties, string]> = [
	['title', 'dc:title'],
	['subject', 'dc:subject'],
	['creator', 'dc:creator'],
	['keywords', 'cp:keywords'],
	['description', 'dc:description'],
	['lastModifiedBy', 'cp:lastModifiedBy'],
	['created', 'dcterms:created'],
	['modified', 'dcterms:modified'],
	['category', 'cp:category'],
];

function setField(root: XmlElement, qname: string, node: XmlElement | null): void {
	const local = localName(qname);
	const i = root.children.findIndex(
		(c) => typeof c === 'object' && 'name' in c && localName(c.name) === local,
	);
	if (i >= 0) {
		if (node) {
			root.children[i] = { ...node, name: (root.children[i] as XmlElement).name };
		} else {
			root.children.splice(i, 1);
		}
		return;
	}
	if (!node) {
		return;
	}
	const order = CORE_FIELDS.map(([, n]) => localName(n));
	const rank = order.indexOf(local);
	const at = root.children.findIndex(
		(c) => typeof c === 'object' && 'name' in c && order.indexOf(localName(c.name)) > rank,
	);
	root.children.splice(at === -1 ? root.children.length : at, 0, node);
}

/** Serialize core properties; `original` keeps fields the model does not cover. */
export function buildCoreXml(
	props: DocumentProperties,
	original: XmlElement | undefined,
	now: Date,
): string {
	const root = original
		? cloneElement(original)
		: el('cp:coreProperties', {
				'xmlns:cp': NS.coreProps,
				'xmlns:dc': NS.dc,
				'xmlns:dcterms': NS.dcterms,
				'xmlns:dcmitype': NS.dcmitype,
				'xmlns:xsi': NS.xsi,
			});
	for (const ns of [
		['xmlns:dc', NS.dc],
		['xmlns:dcterms', NS.dcterms],
		['xmlns:xsi', NS.xsi],
	] as const) {
		if (root.attrs[ns[0]] === undefined) {
			root.attrs[ns[0]] = ns[1];
		}
	}
	for (const [key, qname] of CORE_FIELDS) {
		if (key === 'modified') {
			setField(root, qname, textEl(qname, w3cdtf(now), { 'xsi:type': 'dcterms:W3CDTF' }));
			continue;
		}
		const value = props[key];
		if (key === 'created') {
			if (typeof value === 'string' && value) {
				setField(root, qname, textEl(qname, w3cdtf(value), { 'xsi:type': 'dcterms:W3CDTF' }));
			}
			continue;
		}
		setField(root, qname, typeof value === 'string' && value !== '' ? textEl(qname, value) : null);
	}
	return serializeXml(root);
}

// ── app.xml ───────────────────────────────────────────────────────────────

function vector(parent: XmlElement | undefined): XmlElement | undefined {
	return elements(parent).find((c) => localName(c.name) === 'vector');
}

function lpstrVector(values: readonly string[]): XmlElement {
	return el(
		'vt:vector',
		{ size: values.length, baseType: 'lpstr' },
		values.map((v) => textEl('vt:lpstr', v)),
	);
}

function pairsVector(groups: ReadonlyArray<[string, number]>): XmlElement {
	return el(
		'vt:vector',
		{ size: groups.length * 2, baseType: 'variant' },
		groups.flatMap(([name, count]) => [
			el('vt:variant', {}, [textEl('vt:lpstr', name)]),
			el('vt:variant', {}, [textEl('vt:i4', String(count))]),
		]),
	);
}

function setAppChild(root: XmlElement, name: string, value: XmlElement): void {
	const i = root.children.findIndex(
		(c) => typeof c === 'object' && 'name' in c && localName(c.name) === name,
	);
	if (i >= 0) {
		root.children[i] = value;
	} else {
		root.children.push(value);
	}
}

/** Replace the worksheet group of HeadingPairs / TitlesOfParts with the current sheet names. */
function updateTitles(
	root: XmlElement,
	previous: readonly string[],
	names: readonly string[],
): void {
	const variants = elements(
		vector(elements(root).find((c) => localName(c.name) === 'HeadingPairs')),
	);
	const groups: Array<[string, number]> = [];
	for (let i = 0; i + 1 < variants.length; i += 2) {
		groups.push([textContent(variants[i]), Number(textContent(variants[i + 1])) || 0]);
	}
	const titles = elements(
		vector(elements(root).find((c) => localName(c.name) === 'TitlesOfParts')),
	).map((t) => textContent(t));
	let nextGroups: Array<[string, number]> = [['Worksheets', names.length]];
	let nextTitles: string[] = [...names];
	if (groups.length && groups.reduce((n, g) => n + g[1], 0) === titles.length) {
		// The worksheet group is the one listing the previous sheet names (else the first group).
		let target = 0;
		let start = 0;
		for (let gi = 0, offset = 0; gi < groups.length; offset += groups[gi][1], gi++) {
			if (titles.slice(offset, offset + groups[gi][1]).join('\u0000') === previous.join('\u0000')) {
				target = gi;
				start = offset;
				break;
			}
		}
		nextGroups = groups.map((g, i) => (i === target ? [g[0], names.length] : g));
		nextTitles = [...titles.slice(0, start), ...names, ...titles.slice(start + groups[target][1])];
	}
	setAppChild(root, 'HeadingPairs', el('HeadingPairs', {}, [pairsVector(nextGroups)]));
	setAppChild(root, 'TitlesOfParts', el('TitlesOfParts', {}, [lpstrVector(nextTitles)]));
}

export interface AppXmlInput {
	sheetNames: string[];
	previousSheetNames: string[];
	application: string;
	appVersion?: string;
	company?: string;
	manager?: string;
}

/** Serialize extended properties (a fresh part, or the original with updated sheet titles). */
export function buildAppXml(input: AppXmlInput, original: XmlElement | undefined): string {
	if (original) {
		const root = cloneElement(original);
		updateTitles(root, input.previousSheetNames, input.sheetNames);
		// Extended properties are an xsd:all group, so replacing or appending keeps the part valid.
		for (const [name, value] of [
			['Company', input.company],
			['Manager', input.manager],
		] as const) {
			if (value === undefined) {
				continue;
			}
			if (value === '' && name === 'Manager') {
				root.children = root.children.filter(
					(c) => !(typeof c === 'object' && 'name' in c && localName(c.name) === name),
				);
			} else {
				setAppChild(root, name, textEl(name, value));
			}
		}
		return serializeXml(root);
	}
	const version =
		input.appVersion && /^\d{1,2}\.\d{4}$/.test(input.appVersion) ? input.appVersion : '16.0300';
	return serializeXml(
		el('Properties', { xmlns: NS.extendedProps, 'xmlns:vt': NS.vt }, [
			textEl('Application', input.application),
			textEl('DocSecurity', '0'),
			textEl('ScaleCrop', 'false'),
			el('HeadingPairs', {}, [pairsVector([['Worksheets', input.sheetNames.length]])]),
			el('TitlesOfParts', {}, [lpstrVector(input.sheetNames)]),
			input.manager ? textEl('Manager', input.manager) : undefined,
			textEl('Company', input.company ?? ''),
			textEl('LinksUpToDate', 'false'),
			textEl('SharedDoc', 'false'),
			textEl('HyperlinksChanged', 'false'),
			textEl('AppVersion', version),
		]),
	);
}

// ── custom.xml ────────────────────────────────────────────────────────────

const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const VT_TYPES = new Set([
	'lpwstr',
	'lpstr',
	'i4',
	'i8',
	'int',
	'r4',
	'r8',
	'bool',
	'filetime',
	'decimal',
]);

type CustomProp = NonNullable<DocumentProperties['custom']>[number];

function vtType(p: CustomProp): string {
	if (p.type && VT_TYPES.has(p.type)) {
		return p.type;
	}
	if (typeof p.value === 'boolean') {
		return 'bool';
	}
	if (typeof p.value === 'number') {
		return Number.isInteger(p.value) && Math.abs(p.value) < 2 ** 31 ? 'i4' : 'r8';
	}
	return 'lpwstr';
}

function vtValue(p: CustomProp): string {
	return typeof p.value === 'boolean' ? (p.value ? 'true' : 'false') : String(p.value);
}

/** Serialize custom document properties. */
export function buildCustomXml(props: readonly CustomProp[]): string {
	return serializeXml(
		el(
			'Properties',
			{ xmlns: NS.customProps, 'xmlns:vt': NS.vt },
			props.map((p, i) =>
				el('property', { fmtid: FMTID, pid: i + 2, name: p.name }, [
					textEl(`vt:${vtType(p)}`, vtValue(p)),
				]),
			),
		),
	);
}

/** The model's custom properties differ from the original part. */
export function customChanged(props: readonly CustomProp[], originalXml: string): boolean {
	try {
		const root = parseXml(originalXml).root;
		const orig = elements(root).map((p) => {
			const v = elements(p)[0];
			return `${p.attrs.name}\u0000${v ? localName(v.name) : ''}\u0000${textContent(v)}`;
		});
		const model = props.map((p) => `${p.name}\u0000${vtType(p)}\u0000${vtValue(p)}`);
		return orig.join('\u0001') !== model.join('\u0001');
	} catch {
		return true;
	}
}

/**
 * `xl/styles.xml`. Collections are append-only: original entries are
 * re-emitted from their raw XML and only entries beyond the original
 * lengths are serialized from the model. A loaded stylesheet that did not
 * grow is left untouched.
 */
import type { StyleSheet, StyleSheetRaw } from '../types/style';
import { cloneElement, el, localName, namePrefix, parseXml } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement, XmlNode } from '../xml/types';
import {
	borderXml,
	cellStyleXml,
	dxfXml,
	fillXml,
	fontXml,
	numFmtXml,
	tableStyleXml,
	xfXml,
} from './style-entries';
import {
	findChild,
	findChildren,
	mergeAttrs,
	normalizeArgb,
	normalizeRoot,
	serializeRootWith,
	setOrderedChild,
} from './xml-utils';

export const STYLESHEET_ORDER = [
	'numFmts',
	'fonts',
	'fills',
	'borders',
	'cellStyleXfs',
	'cellXfs',
	'cellStyles',
	'dxfs',
	'tableStyles',
	'colors',
	'extLst',
] as const;

type Collection =
	| 'fonts'
	| 'fills'
	| 'borders'
	| 'cellStyleXfs'
	| 'cellXfs'
	| 'cellStyles'
	| 'dxfs';

const ENTRY: Record<Collection, string> = {
	fonts: 'font',
	fills: 'fill',
	borders: 'border',
	cellStyleXfs: 'xf',
	cellXfs: 'xf',
	cellStyles: 'cellStyle',
	dxfs: 'dxf',
};

/** Original entries: `styles.raw` when the reader kept them, else parsed from the root. */
export function rawEntries(
	raw: StyleSheetRaw | undefined,
	root: XmlElement | undefined,
	p = '',
): Required<Omit<StyleSheetRaw, 'root'>> {
	const pick = (c: Collection): XmlElement[] =>
		raw?.[c] ?? findChildren(findChild(root, c, p), ENTRY[c], p);
	return {
		fonts: pick('fonts'),
		fills: pick('fills'),
		borders: pick('borders'),
		cellStyleXfs: pick('cellStyleXfs'),
		cellXfs: pick('cellXfs'),
		cellStyles: pick('cellStyles'),
		dxfs: pick('dxfs'),
	};
}

function numFmtsOf(root: XmlElement | undefined, p: string): Record<number, string> {
	const out: Record<number, string> = {};
	for (const nf of findChildren(findChild(root, 'numFmts', p), 'numFmt', p)) {
		out[Number(nf.attrs.numFmtId)] = nf.attrs.formatCode ?? '';
	}
	return out;
}

function sameNumFmts(a: Record<number, string>, b: Record<number, string>): boolean {
	const ka = Object.keys(a);
	return ka.length === Object.keys(b).length && ka.every((k) => b[Number(k)] === a[Number(k)]);
}

function indexedOf(root: XmlElement | undefined, p: string): string[] {
	const colors = findChild(root, 'colors', p);
	return findChildren(findChild(colors, 'indexedColors', p), 'rgbColor', p).map(
		(c) => c.attrs.rgb ?? '',
	);
}

/** The stylesheet changed compared to its source (something was appended or replaced). */
export function stylesChanged(styles: StyleSheet, root: XmlElement | undefined): boolean {
	if (!root) {
		return true;
	}
	const p = namePrefix(root.name);
	const raw = rawEntries(styles.raw, root, p);
	for (const c of Object.keys(ENTRY) as Collection[]) {
		if (styles[c].length !== raw[c].length) {
			return true;
		}
	}
	if (!sameNumFmts(styles.numFmts, numFmtsOf(root, p))) {
		return true;
	}
	const ts = findChild(root, 'tableStyles', p);
	if ((styles.customTableStyles?.length ?? 0) !== findChildren(ts, 'tableStyle', p).length) {
		return true;
	}
	if (
		(styles.defaultTableStyle !== undefined &&
			styles.defaultTableStyle !== ts?.attrs.defaultTableStyle) ||
		(styles.defaultPivotStyle !== undefined &&
			styles.defaultPivotStyle !== ts?.attrs.defaultPivotStyle)
	) {
		return true;
	}
	const indexed = indexedOf(root, p);
	const model = styles.indexedColors ?? [];
	return (
		indexed.length !== model.length ||
		model.some((c, i) => normalizeArgb(c) !== normalizeArgb(indexed[i]))
	);
}

function collection<T>(
	name: string,
	orig: XmlElement | undefined,
	items: readonly T[],
	raw: readonly XmlElement[],
	write: (item: T) => string,
): XmlElement {
	const kids: XmlNode[] = [];
	let fresh = '';
	items.forEach((item, i) => {
		if (i < raw.length) {
			kids.push(raw[i]);
		} else {
			fresh += write(item);
		}
	});
	if (fresh) {
		kids.push({ raw: fresh });
	}
	return { name, attrs: mergeAttrs(orig?.attrs, { count: items.length }), children: kids };
}

/** Serialize the stylesheet (from the original root when there is one). */
export function buildStylesXml(styles: StyleSheet, sourceRoot: XmlElement | undefined): string {
	const root = sourceRoot ? cloneElement(sourceRoot) : el('styleSheet', { xmlns: NS.main });
	const p = normalizeRoot(root, 'styleSheet', NS.main);
	const raw = rawEntries(styles.raw, sourceRoot, p);
	const set = (name: string, value: XmlElement | null): void =>
		setOrderedChild(root, name, value, STYLESHEET_ORDER, p);
	const find = (name: string): XmlElement | undefined => findChild(root, name, p);

	const ids = Object.keys(styles.numFmts)
		.map(Number)
		.filter((id) => Number.isInteger(id) && id >= 0)
		.sort((a, b) => a - b);
	set(
		'numFmts',
		ids.length
			? {
					name: 'numFmts',
					attrs: mergeAttrs(find('numFmts')?.attrs, { count: ids.length }),
					children: [{ raw: ids.map((id) => numFmtXml(id, styles.numFmts[id])).join('') }],
				}
			: null,
	);
	set(
		'fonts',
		collection('fonts', find('fonts'), styles.fonts, raw.fonts, (f) => fontXml(f)),
	);
	set(
		'fills',
		collection('fills', find('fills'), styles.fills, raw.fills, (f) => fillXml(f)),
	);
	set(
		'borders',
		collection('borders', find('borders'), styles.borders, raw.borders, (b) => borderXml(b)),
	);
	set(
		'cellStyleXfs',
		collection('cellStyleXfs', find('cellStyleXfs'), styles.cellStyleXfs, raw.cellStyleXfs, (x) =>
			xfXml(x, false),
		),
	);
	set(
		'cellXfs',
		collection('cellXfs', find('cellXfs'), styles.cellXfs, raw.cellXfs, (x) => xfXml(x, true)),
	);
	set(
		'cellStyles',
		collection('cellStyles', find('cellStyles'), styles.cellStyles, raw.cellStyles, cellStyleXml),
	);
	set('dxfs', collection('dxfs', find('dxfs'), styles.dxfs, raw.dxfs, dxfXml));

	const origTs = find('tableStyles');
	const custom = styles.customTableStyles ?? [];
	const rawTs = findChildren(origTs, 'tableStyle', p);
	set('tableStyles', {
		name: 'tableStyles',
		attrs: mergeAttrs(origTs?.attrs, {
			count: custom.length,
			defaultTableStyle:
				styles.defaultTableStyle ?? origTs?.attrs.defaultTableStyle ?? 'TableStyleMedium2',
			defaultPivotStyle:
				styles.defaultPivotStyle ?? origTs?.attrs.defaultPivotStyle ?? 'PivotStyleLight16',
		}),
		children: custom.map((t, i) => (i < rawTs.length ? rawTs[i] : { raw: tableStyleXml(t) })),
	});

	const origColors = find('colors');
	const mru = findChild(origColors, 'mruColors', p);
	const indexed = styles.indexedColors?.length
		? el(
				'indexedColors',
				{},
				styles.indexedColors.map((c) => el('rgbColor', { rgb: normalizeArgb(c) })),
			)
		: null;
	set('colors', indexed || mru ? el('colors', origColors?.attrs ?? {}, [indexed, mru]) : null);
	return serializeRootWith(root);
}

/** Parse a styles part root (used when the model carries no raw XML). */
export function parseStylesRoot(xml: string): XmlElement | undefined {
	try {
		const root = parseXml(xml).root;
		return localName(root.name) === 'styleSheet' ? root : undefined;
	} catch {
		return undefined;
	}
}

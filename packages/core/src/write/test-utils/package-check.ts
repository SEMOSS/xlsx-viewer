/**
 * Test helpers: structural validation of a written package (the checks that
 * make Excel offer to repair a file) and XML access shortcuts.
 */
import { lettersToColumn } from '../../address/cell-address';
import { OpcPackage, partExtension, relsPathFor, resolveRelTarget } from '../../package';
import { elements, isElement, localName, parseXml } from '../../xml';
import type { XmlElement } from '../../xml/types';
import { STYLESHEET_ORDER } from '../styles-xml';
import { WORKBOOK_ORDER } from '../workbook-xml';
import { WORKSHEET_ORDER } from '../worksheet-xml';

export async function openPackage(bytes: Uint8Array): Promise<OpcPackage> {
	return OpcPackage.open(bytes);
}

/** Parsed root of a part (throws when missing). */
export async function partRoot(pkg: OpcPackage, path: string): Promise<XmlElement> {
	const text = await pkg.text(path);
	if (text === undefined) {
		throw new Error(`missing part ${path}`);
	}
	return parseXml(text).root;
}

/** Local names of the element children. */
export function childNames(el: XmlElement | undefined): string[] {
	return elements(el).map((c) => localName(c.name));
}

/** First descendant path by local names. */
export function pick(el: XmlElement | undefined, ...names: string[]): XmlElement | undefined {
	let cur = el;
	for (const n of names) {
		cur = elements(cur).find((c) => localName(c.name) === n);
	}
	return cur;
}

export function pickAll(el: XmlElement | undefined, name: string): XmlElement[] {
	return elements(el).filter((c) => localName(c.name) === name);
}

function checkOrder(
	problems: string[],
	part: string,
	root: XmlElement,
	order: readonly string[],
): void {
	let last = -1;
	for (const c of root.children) {
		if (!isElement(c) || c.name.includes(':')) {
			continue;
		}
		const r = order.indexOf(c.name);
		if (r === -1) {
			problems.push(`${part}: unknown child <${c.name}>`);
			continue;
		}
		if (r < last) {
			problems.push(`${part}: <${c.name}> out of schema order`);
		}
		last = r;
	}
}

function refCol(ref: string): number {
	return lettersToColumn(ref.replace(/\d+$/, ''));
}

function checkSheet(
	problems: string[],
	part: string,
	root: XmlElement,
	styleCount: number,
	sstCount: number,
): void {
	checkOrder(problems, part, root, WORKSHEET_ORDER);
	let lastRow = 0;
	for (const row of pickAll(pick(root, 'sheetData'), 'row')) {
		const r = Number(row.attrs.r);
		if (!(r > lastRow)) {
			problems.push(`${part}: row ${r} not ascending`);
		}
		lastRow = r;
		let lastCol = -1;
		for (const c of pickAll(row, 'c')) {
			const ref = c.attrs.r ?? '';
			if (!ref.endsWith(String(r))) {
				problems.push(`${part}: cell ${ref} outside row ${r}`);
			}
			const col = refCol(ref);
			if (col <= lastCol) {
				problems.push(`${part}: cell ${ref} not ascending`);
			}
			lastCol = col;
			if (c.attrs.s !== undefined && Number(c.attrs.s) >= styleCount) {
				problems.push(`${part}: cell ${ref} style ${c.attrs.s} >= ${styleCount}`);
			}
			if (c.attrs.t === 's') {
				const v = Number(pick(c, 'v')?.children[0]);
				if (!(v >= 0 && v < sstCount)) {
					problems.push(`${part}: cell ${ref} shared string ${v} out of range`);
				}
			}
		}
	}
}

/** Problems that would make Excel repair the file (empty when the package looks valid). */
export async function validatePackage(pkg: OpcPackage): Promise<string[]> {
	const problems: string[] = [];
	const parts = pkg.partPaths().filter((p) => p !== '[Content_Types].xml');
	const lower = new Set(parts.map((p) => p.toLowerCase()));
	const roots = new Map<string, XmlElement>();
	for (const p of parts) {
		const ext = partExtension(p);
		if (!p.endsWith('.rels') && !pkg.contentTypes.get(p)) {
			problems.push(`${p}: no content type`);
		}
		if (ext === 'xml' || ext === 'rels' || ext === 'vml') {
			try {
				roots.set(p, parseXml((await pkg.text(p)) ?? '').root);
			} catch (e) {
				problems.push(`${p}: not well-formed (${(e as Error).message})`);
			}
		}
	}
	for (const o of pkg.contentTypes.overrides.keys()) {
		if (!lower.has(o.toLowerCase())) {
			problems.push(`override for missing part ${o}`);
		}
	}
	for (const p of ['', ...parts.filter((x) => !x.endsWith('.rels'))]) {
		const rels = await pkg.rels(p);
		const ids = new Set<string>();
		for (const r of rels) {
			if (ids.has(r.id)) {
				problems.push(`${p || '/'}: duplicate rel id ${r.id}`);
			}
			ids.add(r.id);
			if (r.targetMode !== 'External' && !r.target.startsWith('#')) {
				const t = resolveRelTarget(p, r.target);
				if (!lower.has(t.toLowerCase())) {
					problems.push(`${p || '/'}: rel ${r.id} -> missing ${t}`);
				}
			}
		}
		const ext = partExtension(p);
		const text = p && (ext === 'xml' || ext === 'vml') ? await pkg.text(p) : undefined;
		if (text) {
			for (const m of text.matchAll(/\br:(?:id|embed|link|pict)="([^"]*)"/g)) {
				if (!ids.has(m[1])) {
					problems.push(`${p}: r:id ${m[1]} has no relationship`);
				}
			}
		}
		if (p && rels.length && !lower.has(relsPathFor(p).toLowerCase())) {
			problems.push(`${p}: rels part missing`);
		}
	}
	const wbPath =
		parts.find((p) => pkg.contentTypes.get(p)?.includes('sheet.main+xml')) ?? 'xl/workbook.xml';
	const wb = roots.get(wbPath);
	const styles = [...roots.entries()].find(([p]) =>
		pkg.contentTypes.get(p)?.endsWith('styles+xml'),
	)?.[1];
	const sst = [...roots.entries()].find(([p]) =>
		pkg.contentTypes.get(p)?.endsWith('sharedStrings+xml'),
	)?.[1];
	const styleCount = pickAll(pick(styles, 'cellXfs'), 'xf').length;
	const sstCount = pickAll(sst, 'si').length;
	if (styles) {
		checkOrder(problems, 'styles', styles, STYLESHEET_ORDER);
		for (const name of [
			'fonts',
			'fills',
			'borders',
			'cellStyleXfs',
			'cellXfs',
			'cellStyles',
			'dxfs',
		]) {
			const c = pick(styles, name);
			if (c && Number(c.attrs.count) !== elements(c).length) {
				problems.push(`styles: ${name} count mismatch`);
			}
		}
	}
	if (wb) {
		checkOrder(problems, wbPath, wb, WORKBOOK_ORDER);
		const sheets = pickAll(pick(wb, 'sheets'), 'sheet');
		const names = new Set(sheets.map((s) => s.attrs.name?.toLowerCase()));
		const sheetIds = new Set(sheets.map((s) => s.attrs.sheetId));
		if (names.size !== sheets.length || sheetIds.size !== sheets.length) {
			problems.push('workbook: duplicate sheet names or ids');
		}
		if (!sheets.some((s) => !s.attrs.state || s.attrs.state === 'visible')) {
			problems.push('workbook: no visible sheet');
		}
		for (const d of pickAll(pick(wb, 'definedNames'), 'definedName')) {
			if (d.attrs.localSheetId !== undefined && Number(d.attrs.localSheetId) >= sheets.length) {
				problems.push(`workbook: defined name ${d.attrs.name} scope out of range`);
			}
		}
	}
	const tableIds = new Set<string>();
	const tableNames = new Set<string>();
	for (const [p, root] of roots) {
		const ct = pkg.contentTypes.get(p) ?? '';
		if (ct.endsWith('worksheet+xml')) {
			checkSheet(problems, p, root, styleCount, sstCount);
		} else if (ct.endsWith('table+xml')) {
			const id = root.attrs.id;
			const name = root.attrs.name?.toLowerCase() ?? '';
			if (tableIds.has(id) || tableNames.has(name)) {
				problems.push(`${p}: duplicate table id or name`);
			}
			tableIds.add(id);
			tableNames.add(name);
			const cols = pickAll(pick(root, 'tableColumns'), 'tableColumn');
			const [a, b] = (root.attrs.ref ?? '').split(':');
			if (b && refCol(b) - refCol(a) + 1 !== cols.length) {
				problems.push(`${p}: column count does not match ref`);
			}
		}
	}
	return problems;
}

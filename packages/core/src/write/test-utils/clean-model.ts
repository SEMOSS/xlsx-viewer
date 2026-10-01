/**
 * A minimal "unchanged" Workbook for a source package, built without the
 * reader: every dirty flag is false and the collections that drive change
 * detection (sheet list, style counts, number formats, defined names) match
 * the file. Used to check that saving an unchanged model is byte-identical.
 */
import { findRelByType, type OpcPackage, resolveRelTarget } from '../../package';
import { createWorksheet } from '../../sheet/worksheet';
import { createDefaultTheme } from '../../styles/defaults';
import type { Worksheet } from '../../types/sheet';
import type { CellXf, StyleSheet } from '../../types/style';
import type { DefinedName, Workbook } from '../../types/workbook';
import { textContent } from '../../xml';
import { REL } from '../../xml/namespaces';
import { partRoot, pick, pickAll } from './package-check';

async function styleSheetOf(pkg: OpcPackage, path: string | undefined): Promise<StyleSheet> {
	const root = path ? await partRoot(pkg, path) : undefined;
	const count = (name: string, entry: string): number => pickAll(pick(root, name), entry).length;
	const numFmts: Record<number, string> = {};
	for (const nf of pickAll(pick(root, 'numFmts'), 'numFmt')) {
		numFmts[Number(nf.attrs.numFmtId)] = nf.attrs.formatCode ?? '';
	}
	const xf = (): CellXf => ({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
	const ts = pick(root, 'tableStyles');
	return {
		numFmts,
		fonts: Array.from({ length: count('fonts', 'font') }, () => ({})),
		fills: Array.from({ length: count('fills', 'fill') }, () => ({
			type: 'pattern' as const,
			patternType: 'none' as const,
		})),
		borders: Array.from({ length: count('borders', 'border') }, () => ({})),
		cellStyleXfs: Array.from({ length: count('cellStyleXfs', 'xf') }, xf),
		cellXfs: Array.from({ length: count('cellXfs', 'xf') }, xf),
		cellStyles: Array.from({ length: count('cellStyles', 'cellStyle') }, () => ({
			name: 'Normal',
			xfId: 0,
		})),
		dxfs: Array.from({ length: count('dxfs', 'dxf') }, () => ({})),
		defaultTableStyle: ts?.attrs.defaultTableStyle,
		defaultPivotStyle: ts?.attrs.defaultPivotStyle,
		customTableStyles: pickAll(ts, 'tableStyle').map((t) => ({
			name: t.attrs.name ?? '',
			elements: [],
		})),
	};
}

/** Build the clean model of a package. */
export async function cleanModel(pkg: OpcPackage): Promise<Workbook> {
	const main = findRelByType(await pkg.rels(''), REL.officeDocument)!;
	const wbPath = resolveRelTarget('', main.target);
	const wbRoot = await partRoot(pkg, wbPath);
	const wbRels = await pkg.rels(wbPath);
	const path = (type: string): string | undefined => {
		const r = findRelByType(wbRels, type);
		return r ? resolveRelTarget(wbPath, r.target) : undefined;
	};
	const sheets: Worksheet[] = [];
	for (const s of pickAll(pick(wbRoot, 'sheets'), 'sheet')) {
		const rel = wbRels.find((r) => r.id === s.attrs['r:id'])!;
		const partPath = resolveRelTarget(wbPath, rel.target);
		const ws = createWorksheet(s.attrs.name ?? '', Number(s.attrs.sheetId));
		ws.state = (s.attrs.state as Worksheet['state']) ?? 'visible';
		ws.dirty = false;
		ws.drawingsDirty = false;
		ws.commentsDirty = false;
		ws.source = { partPath, rels: await pkg.rels(partPath) };
		sheets.push(ws);
	}
	const definedNames: DefinedName[] = pickAll(pick(wbRoot, 'definedNames'), 'definedName').map(
		(d) => ({
			name: d.attrs.name ?? '',
			formula: textContent(d),
			localSheet: d.attrs.localSheetId !== undefined ? Number(d.attrs.localSheetId) : undefined,
			hidden: d.attrs.hidden === '1' ? true : undefined,
		}),
	);
	const calcPr = pick(wbRoot, 'calcPr');
	const themePath = path(REL.theme);
	const themeXml = themePath ? await pkg.text(themePath) : undefined;
	return {
		sheets,
		definedNames,
		styles: await styleSheetOf(pkg, path(REL.styles)),
		theme: { ...createDefaultTheme(), rawXml: themeXml },
		properties: {},
		view: { activeTab: Number(pick(wbRoot, 'bookViews', 'workbookView')?.attrs.activeTab ?? 0) },
		date1904: pick(wbRoot, 'workbookPr')?.attrs.date1904 === '1',
		calc: calcPr?.attrs.calcId !== undefined ? { calcId: Number(calcPr.attrs.calcId) } : {},
		media: new Map(),
		persons: [],
		warnings: [],
		sourcePackage: pkg,
		dirty: false,
	};
}

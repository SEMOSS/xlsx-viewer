/**
 * `xl/workbook.xml`, regenerated from the original root when there is one
 * (file version, extensions, pivot caches and other unknown children stay).
 */
import { formatRange } from '../address/range';
import { quoteSheetName } from '../address/sheet-name';
import type { Worksheet } from '../types/sheet';
import type { DefinedName, Workbook } from '../types/workbook';
import { cloneElement, el, textContent, textEl } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { fileFormula } from './formulas';
import {
	ensureNamespace,
	findChild,
	findChildren,
	flag,
	mergeAttrs,
	normalizeRoot,
	setOrderedChild,
	unflag,
} from './xml-utils';

export const WORKBOOK_ORDER = [
	'fileVersion',
	'fileSharing',
	'workbookPr',
	'workbookProtection',
	'bookViews',
	'sheets',
	'functionGroups',
	'externalReferences',
	'definedNames',
	'calcPr',
	'oleSize',
	'customWorkbookViews',
	'pivotCaches',
	'smartTagPr',
	'smartTagTypes',
	'webPublishing',
	'fileRecoveryPr',
	'webPublishObjects',
	'extLst',
] as const;

export interface WorkbookSheetEntry {
	sheet: Worksheet;
	relId: string;
	/** Name and sheetId as written (validated / deduplicated). */
	name: string;
	sheetId: number;
}

export interface WorkbookXmlInput {
	workbook: Workbook;
	original?: XmlElement;
	sheets: WorkbookSheetEntry[];
	/** Names of the original sheets by original index (to match defined-name extras). */
	originalSheetNames: string[];
	/** Force `fullCalcOnLoad` (stale or never calculated formulas were written). */
	fullCalcOnLoad: boolean;
}

/** Defined names as written: valid scope, no duplicates, plus missing filter-database names. */
export function effectiveDefinedNames(workbook: Workbook): DefinedName[] {
	const out: DefinedName[] = [];
	const seen = new Set<string>();
	for (const n of workbook.definedNames) {
		if (!n.name || n.formula === undefined) {
			continue;
		}
		if (
			n.localSheet !== undefined &&
			(n.localSheet < 0 || n.localSheet >= workbook.sheets.length)
		) {
			continue;
		}
		const key = `${n.name.toLowerCase()}\u0000${n.localSheet ?? ''}`;
		if (seen.has(key)) {
			continue;
		}
		seen.add(key);
		out.push(n);
	}
	workbook.sheets.forEach((s, i) => {
		const key = `_xlnm._filterdatabase\u0000${i}`;
		if (s.autoFilter && !seen.has(key)) {
			seen.add(key);
			out.push({
				name: '_xlnm._FilterDatabase',
				localSheet: i,
				hidden: true,
				builtin: true,
				formula: `${quoteSheetName(s.name)}!${formatRange(s.autoFilter.ref, { absolute: true, forceCells: true })}`,
			});
		}
	});
	return out;
}

/** The active tab as written: clamped to the sheet list and moved off hidden sheets. */
export function activeTab(workbook: Workbook): number {
	const sheets = workbook.sheets;
	let tab = Math.min(
		Math.max(0, Math.trunc(workbook.view.activeTab || 0)),
		Math.max(0, sheets.length - 1),
	);
	if (sheets[tab] && sheets[tab].state !== 'visible') {
		const visible = sheets.findIndex((s) => s.state === 'visible');
		tab = visible >= 0 ? visible : tab;
	}
	return tab;
}

function definedNamesElement(
	input: WorkbookXmlInput,
	orig: XmlElement | undefined,
): XmlElement | null {
	const names = effectiveDefinedNames(input.workbook);
	if (!names.length) {
		return null;
	}
	const extras = new Map<string, XmlElement>();
	for (const d of findChildren(orig, 'definedName')) {
		const scope =
			d.attrs.localSheetId !== undefined
				? input.originalSheetNames[Number(d.attrs.localSheetId)]
				: '';
		extras.set(`${d.attrs.name?.toLowerCase()}\u0000${scope ?? ''}`, d);
	}
	const kids = names.map((n) => {
		const scope = n.localSheet !== undefined ? input.sheets[n.localSheet]?.sheet.name : '';
		const original = extras.get(`${n.name.toLowerCase()}\u0000${scope ?? ''}`);
		const attrs = mergeAttrs(original?.attrs, {
			name: n.name,
			comment: n.comment,
			localSheetId: n.localSheet,
			hidden: flag(n.hidden),
		});
		const formula = fileFormula(n.formula);
		// Keep the original text when only the prefix normalisation differs.
		const text =
			original && fileFormula(textContent(original)) === formula ? textContent(original) : formula;
		return textEl('definedName', text, attrs);
	});
	return el('definedNames', {}, kids);
}

/** Build the workbook root element. */
export function buildWorkbookElement(input: WorkbookXmlInput): XmlElement {
	const { workbook } = input;
	const root = input.original
		? cloneElement(input.original)
		: el('workbook', { xmlns: NS.main, 'xmlns:r': NS.r });
	const p = normalizeRoot(root, 'workbook', NS.main);
	const r = ensureNamespace(root, 'r', NS.r);
	const set = (name: string, value: XmlElement | null): void =>
		setOrderedChild(root, name, value, WORKBOOK_ORDER, p);
	const find = (name: string): XmlElement | undefined => findChild(root, name, p);

	set(
		'workbookPr',
		el('workbookPr', mergeAttrs(find('workbookPr')?.attrs, { date1904: flag(workbook.date1904) })),
	);
	const prot = workbook.protection;
	set(
		'workbookProtection',
		prot
			? el('workbookProtection', {
					workbookPassword: prot.workbookPassword,
					workbookAlgorithmName: prot.workbookAlgorithmName,
					workbookHashValue: prot.workbookHashValue,
					workbookSaltValue: prot.workbookSaltValue,
					workbookSpinCount: prot.workbookSpinCount,
					lockStructure: flag(prot.lockStructure),
					lockWindows: flag(prot.lockWindows),
				})
			: null,
	);

	const views = find('bookViews') ? cloneElement(find('bookViews')!) : el('bookViews');
	const firstView = findChild(views, 'workbookView', p);
	const view = workbook.view;
	const tab = activeTab(workbook);
	const first = Math.min(
		Math.max(0, view.firstSheet ?? 0),
		Math.max(0, workbook.sheets.length - 1),
	);
	const viewEl = el(
		'workbookView',
		mergeAttrs(
			firstView?.attrs ?? {
				xWindow: '0',
				yWindow: '0',
				windowWidth: '28800',
				windowHeight: '12300',
			},
			{
				showHorizontalScroll: unflag(view.showHorizontalScroll),
				showVerticalScroll: unflag(view.showVerticalScroll),
				showSheetTabs: unflag(view.showSheetTabs),
				tabRatio: view.tabRatio,
				firstSheet: first || undefined,
				activeTab: tab || undefined,
			},
		),
		firstView ? firstView.children : [],
	);
	if (firstView) {
		views.children[views.children.indexOf(firstView)] = viewEl;
	} else {
		views.children.unshift(viewEl);
	}
	set('bookViews', views);

	set(
		'sheets',
		el(
			'sheets',
			{},
			input.sheets.map((s) =>
				el('sheet', {
					name: s.name,
					sheetId: s.sheetId,
					state: s.sheet.state !== 'visible' ? s.sheet.state : undefined,
					[`${r}:id`]: s.relId,
				}),
			),
		),
	);
	set('definedNames', definedNamesElement(input, find('definedNames')));

	const calc = workbook.calc;
	set(
		'calcPr',
		el(
			'calcPr',
			mergeAttrs(find('calcPr')?.attrs, {
				calcId: calc.calcId ?? (input.original ? undefined : 191029),
				calcMode: calc.calcMode && calc.calcMode !== 'auto' ? calc.calcMode : undefined,
				fullCalcOnLoad: calc.fullCalcOnLoad || input.fullCalcOnLoad ? 1 : undefined,
				refMode: calc.refMode === 'R1C1' ? 'R1C1' : undefined,
				iterate: flag(calc.iterate),
				iterateCount: calc.iterate ? calc.iterateCount : undefined,
				iterateDelta: calc.iterate ? calc.iterateDelta : undefined,
			}),
		),
	);
	return root;
}

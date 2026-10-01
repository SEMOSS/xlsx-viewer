import { stripFormulaPrefixes } from '../formula/prefixes';
import type { SheetKind } from '../types/sheet';
import type {
	CalcProperties,
	DefinedName,
	WorkbookProtection,
	WorkbookView,
} from '../types/workbook';
import type { XmlElement } from '../xml/types';
import { relTypeSuffix } from './part-io';
import { bool, compact, int, kid, kids, num, oneOf, ooxmlText, str } from './xml-utils';

/** A `<sheet>` entry of workbook.xml. */
export interface SheetEntry {
	name: string;
	sheetId: number;
	state: 'visible' | 'hidden' | 'veryHidden';
	/** Relationship id of the sheet part. */
	relId?: string;
}

export interface WorkbookPart {
	sheets: SheetEntry[];
	definedNames: DefinedName[];
	date1904: boolean;
	view: WorkbookView;
	calc: CalcProperties;
	protection?: WorkbookProtection;
	/** `<externalReferences>` is present (links to other workbooks). */
	externalReferences: boolean;
	/** `<pivotCaches>` is present. */
	pivotCaches: boolean;
}

/** Sheet kind from the relationship type that targets the sheet part. */
export function sheetKindFromRelType(type: string | undefined): SheetKind {
	switch (relTypeSuffix(type ?? '')) {
		case 'chartsheet':
			return 'chartsheet';
		case 'dialogsheet':
			return 'dialogsheet';
		case 'xlMacrosheet':
		case 'xlIntlMacrosheet':
			return 'macrosheet';
		default:
			return 'worksheet';
	}
}

/** Read the normalised `<workbook>` root. */
export function parseWorkbookPart(root: XmlElement): WorkbookPart {
	const sheets: SheetEntry[] = kids(kid(root, 'sheets'), 'sheet').map((s, i) => {
		const state = oneOf(str(s, 'state'), ['visible', 'hidden', 'veryHidden'] as const) ?? 'visible';
		const entry: SheetEntry = {
			name: str(s, 'name') ?? `Sheet${i + 1}`,
			sheetId: int(s, 'sheetId') ?? i + 1,
			state,
		};
		const relId = str(s, 'r:id');
		if (relId) {
			entry.relId = relId;
		}
		return entry;
	});

	const definedNames: DefinedName[] = kids(kid(root, 'definedNames'), 'definedName').map((d) => {
		const name = str(d, 'name') ?? '';
		const dn: DefinedName = { name, formula: stripFormulaPrefixes(ooxmlText(d).trim()) };
		const local = int(d, 'localSheetId');
		if (local !== undefined && local >= 0 && local < sheets.length) {
			dn.localSheet = local;
		}
		if (bool(d, 'hidden')) {
			dn.hidden = true;
		}
		const comment = str(d, 'comment');
		if (comment) {
			dn.comment = comment;
		}
		if (name.toLowerCase().startsWith('_xlnm.')) {
			dn.builtin = true;
		}
		return dn;
	});

	const wbPr = kid(root, 'workbookPr');
	const bookView = kid(kid(root, 'bookViews'), 'workbookView');
	const view = compact<WorkbookView>({
		activeTab: Math.max(
			0,
			Math.min(int(bookView, 'activeTab') ?? 0, Math.max(0, sheets.length - 1)),
		),
		firstSheet: int(bookView, 'firstSheet'),
		showHorizontalScroll: bool(bookView, 'showHorizontalScroll'),
		showVerticalScroll: bool(bookView, 'showVerticalScroll'),
		showSheetTabs: bool(bookView, 'showSheetTabs'),
		tabRatio: int(bookView, 'tabRatio'),
	});

	const calcPr = kid(root, 'calcPr');
	const calc = compact<CalcProperties>({
		calcId: int(calcPr, 'calcId'),
		calcMode: oneOf(str(calcPr, 'calcMode'), ['auto', 'autoNoTable', 'manual'] as const),
		fullCalcOnLoad: bool(calcPr, 'fullCalcOnLoad'),
		iterate: bool(calcPr, 'iterate'),
		iterateCount: int(calcPr, 'iterateCount'),
		iterateDelta: num(calcPr, 'iterateDelta'),
		refMode: oneOf(str(calcPr, 'refMode'), ['A1', 'R1C1'] as const),
	});

	const result: WorkbookPart = {
		sheets,
		definedNames,
		date1904: bool(wbPr, 'date1904') ?? false,
		view,
		calc,
		externalReferences: kids(kid(root, 'externalReferences'), 'externalReference').length > 0,
		pivotCaches: kids(kid(root, 'pivotCaches'), 'pivotCache').length > 0,
	};

	const prot = kid(root, 'workbookProtection');
	if (prot) {
		result.protection = compact<WorkbookProtection>({
			lockStructure: bool(prot, 'lockStructure'),
			lockWindows: bool(prot, 'lockWindows'),
			workbookPassword: str(prot, 'workbookPassword'),
			workbookAlgorithmName: str(prot, 'workbookAlgorithmName'),
			workbookHashValue: str(prot, 'workbookHashValue'),
			workbookSaltValue: str(prot, 'workbookSaltValue'),
			workbookSpinCount: int(prot, 'workbookSpinCount'),
		});
	}
	return result;
}

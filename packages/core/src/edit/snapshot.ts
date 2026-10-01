import type { Worksheet } from '../types/sheet';

/** Shallow copy of every mutable collection of a sheet (cells are immutable, so sharing them is safe). */
export interface SheetSnapshot {
	sheet: Worksheet;
	fields: Partial<Worksheet>;
}

const SHEET_COLLECTION_KEYS = [
	'cells',
	'rows',
	'cols',
	'merges',
	'hyperlinks',
	'conditionalFormats',
	'dataValidations',
	'tables',
	'drawings',
	'notes',
	'threadedComments',
	'sparklineGroups',
] as const;

const SHEET_SCALAR_KEYS = [
	'name',
	'state',
	'view',
	'tabColor',
	'autoFilter',
	'sortState',
	'protection',
	'defaultRowHeight',
	'defaultColWidth',
	'pageSetup',
	'pageMargins',
	'printOptions',
	'headerFooter',
	'rowBreaks',
	'colBreaks',
	'dirty',
	'drawingsDirty',
	'commentsDirty',
	'outlineLevelRow',
	'outlineLevelCol',
] as const;

export function snapshotSheet(sheet: Worksheet): SheetSnapshot {
	const fields: Partial<Worksheet> = {};
	const f = fields as Record<string, unknown>;
	const s = sheet as unknown as Record<string, unknown>;
	f.cells = new Map(sheet.cells);
	f.rows = new Map(sheet.rows);
	for (const key of SHEET_COLLECTION_KEYS) {
		if (key === 'cells' || key === 'rows') {
			continue;
		}
		f[key] = [...(s[key] as unknown[])];
	}
	for (const key of SHEET_SCALAR_KEYS) {
		const v = s[key];
		f[key] = v && typeof v === 'object' ? structuredCloneSafe(v) : v;
	}
	return { sheet, fields };
}

function structuredCloneSafe<T>(v: T): T {
	try {
		return structuredClone(v);
	} catch {
		return JSON.parse(JSON.stringify(v)) as T;
	}
}

export function restoreSheet(snap: SheetSnapshot): void {
	const s = snap.sheet as unknown as Record<string, unknown>;
	for (const [k, v] of Object.entries(snap.fields) as Array<[string, unknown]>) {
		if (v instanceof Map) {
			s[k] = new Map(v as Map<unknown, unknown>);
		} else if (Array.isArray(v)) {
			s[k] = [...v];
		} else if (v && typeof v === 'object') {
			s[k] = structuredCloneSafe(v);
		} else {
			s[k] = v;
		}
	}
}

import { createDefaultStyleSheet, createDefaultTheme } from '../styles/defaults';
import type { Workbook } from '../types/workbook';
import type { Worksheet } from '../types/sheet';
import { createWorksheet } from './worksheet';

/** A new in-memory workbook with one empty sheet (`Sheet1`). */
export function createWorkbook(options: { sheetNames?: string[]; title?: string } = {}): Workbook {
	const names = options.sheetNames?.length ? options.sheetNames : ['Sheet1'];
	const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
	return {
		sheets: names.map((n, i) => createWorksheet(n, i + 1)),
		definedNames: [],
		styles: createDefaultStyleSheet(),
		theme: createDefaultTheme(),
		properties: {
			title: options.title,
			created: now,
			modified: now,
			application: 'xlsx-viewer',
		},
		view: { activeTab: 0 },
		date1904: false,
		calc: { calcId: 191029 },
		media: new Map(),
		persons: [],
		warnings: [],
		dirty: true,
	};
}

/** Look up a sheet by name (case-insensitive, like Excel). */
export function findSheet(workbook: Workbook, name: string): Worksheet | undefined {
	const lower = name.toLowerCase();
	return workbook.sheets.find((s) => s.name.toLowerCase() === lower);
}

/** Index of a sheet by name (case-insensitive), or -1. */
export function findSheetIndex(workbook: Workbook, name: string): number {
	const lower = name.toLowerCase();
	return workbook.sheets.findIndex((s) => s.name.toLowerCase() === lower);
}

/** A sheet name not yet used in the workbook: `Sheet2`, `Sheet3`, ... */
export function nextSheetName(workbook: Workbook, base = 'Sheet'): string {
	for (let i = 1; ; i++) {
		const candidate = `${base}${i}`;
		if (!findSheet(workbook, candidate)) {
			return candidate;
		}
	}
}

/** Next unused OOXML sheetId. */
export function nextSheetId(workbook: Workbook): number {
	return workbook.sheets.reduce((m, s) => Math.max(m, s.sheetId), 0) + 1;
}

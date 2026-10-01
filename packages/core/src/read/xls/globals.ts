/**
 * The workbook globals substream: sheet directory, shared strings,
 * formatting records, defined names, external references and window
 * settings. Each record is read on its own; a damaged record is reported
 * and skipped.
 */
import type { FontSpec } from '../../types/style';
import type { WorkbookProtection } from '../../types/workbook';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import type { SupBook, Xti } from './formula';
import { parseSst, type XlsSharedStrings } from './sst';
import type { XlsWarnings } from './warnings';
import {
	parseFont,
	parseFormat,
	parsePalette,
	parseStyle,
	parseXf,
	type RawStyle,
	type RawXf,
} from './xf-records';

export interface BoundSheet {
	name: string;
	state: 'visible' | 'hidden' | 'veryHidden';
	/** BOUNDSHEET `dt`: 0 worksheet (or dialog sheet), 1 macro sheet, 2 chart, 6 VBA module. */
	type: number;
	/** Stream offset of the sheet's BOF record. */
	offset: number;
}

/** A NAME (Lbl) record. */
export interface RawName {
	/** Name as written in formulas (`_xlnm.Print_Area` for built-in names). */
	name: string;
	builtin: boolean;
	hidden: boolean;
	/** Macro or user-defined function name (fFunc / fProc). */
	func: boolean;
	/** 1-based sheet (BOUNDSHEET order) for sheet-scoped names, 0 for workbook scope. */
	itab: number;
	rgce: Uint8Array;
	extra: Uint8Array;
}

export interface WindowSettings {
	activeTab: number;
	firstTab: number;
	flags: number;
	tabRatio: number;
}

export interface XlsGlobals {
	date1904: boolean;
	sheets: BoundSheet[];
	sst: XlsSharedStrings;
	fonts: FontSpec[];
	formats: Map<number, string>;
	xfs: RawXf[];
	styles: RawStyle[];
	palette?: string[];
	names: RawName[];
	supbooks: SupBook[];
	xti: Xti[];
	window?: WindowSettings;
	lastModifiedBy?: string;
	protection?: WorkbookProtection;
}

/** Built-in name codes; those in {@link XLNM} carry the `_xlnm.` prefix in SpreadsheetML. */
const BUILTIN_NAMES = [
	'Consolidate_Area',
	'Auto_Open',
	'Auto_Close',
	'Extract',
	'Database',
	'Criteria',
	'Print_Area',
	'Print_Titles',
	'Recorder',
	'Data_Form',
	'Auto_Activate',
	'Auto_Deactivate',
	'Sheet_Title',
	'_FilterDatabase',
];
const XLNM = new Set([0, 3, 4, 5, 6, 7, 12, 13]);

const STATES: readonly BoundSheet['state'][] = ['visible', 'hidden', 'veryHidden'];

function parseBoundSheet(record: BiffRecord): BoundSheet {
	const cur = new BiffCursor(record.data, record.breaks);
	const offset = cur.u32();
	const state = STATES[cur.u8() & 0x03] ?? 'hidden';
	const type = cur.u8();
	return { name: cur.shortString(), state, type, offset };
}

function parseName(record: BiffRecord): RawName {
	const cur = new BiffCursor(record.data, record.breaks);
	const flags = cur.u16();
	cur.skip(1);
	const cch = cur.u8();
	const cce = cur.u16();
	cur.skip(2);
	const itab = cur.u16();
	cur.skip(4);
	const raw = cur.stringNoCch(cch);
	const code = raw.length === 1 ? raw.charCodeAt(0) : -1;
	const builtin = (flags & 0x20) !== 0 && BUILTIN_NAMES[code] !== undefined;
	let name = raw;
	if (builtin) {
		name = XLNM.has(code) ? `_xlnm.${BUILTIN_NAMES[code]}` : BUILTIN_NAMES[code];
	}
	const rgce = cur.bytes(Math.min(cce, cur.remaining));
	return {
		name,
		builtin: builtin && XLNM.has(code),
		hidden: (flags & 0x01) !== 0,
		func: (flags & 0x0a) !== 0,
		itab,
		rgce,
		extra: cur.bytes(cur.remaining),
	};
}

/** File name of a SUPBOOK path, with the encoded volume / directory markers turned into `\`. */
function virtPath(path: string): string {
	return path.replace(/^\u0001[\u0001-\u0005]?/, '').replace(/[\u0000-\u0008]/g, '\\');
}

function parseSupBook(record: BiffRecord): SupBook {
	const cur = new BiffCursor(record.data, record.breaks);
	const ctab = cur.u16();
	const cch = cur.u16();
	if (cch === 0x0401) {
		return { kind: 'self', names: [] };
	}
	if (cch === 0x3a01) {
		return { kind: 'addin', names: [] };
	}
	const path = cch > 0 ? cur.stringNoCch(cch) : '';
	if (ctab === 0) {
		return { kind: 'other', names: [] };
	}
	const sheets: string[] = [];
	for (let i = 0; i < ctab && cur.remaining >= 3; i++) {
		sheets.push(cur.unicodeString());
	}
	return { kind: 'external', path: virtPath(path), sheets, names: [] };
}

function parseExternSheet(record: BiffRecord): Xti[] {
	const cur = new BiffCursor(record.data, record.breaks);
	const count = Math.min(cur.u16(), Math.floor(cur.remaining / 6));
	const out: Xti[] = [];
	for (let i = 0; i < count; i++) {
		out.push({ supbook: cur.u16(), first: cur.u16(), last: cur.u16() });
	}
	return out;
}

function parseWindow1(record: BiffRecord): WindowSettings {
	const cur = new BiffCursor(record.data, record.breaks);
	cur.skip(8);
	const flags = cur.u16();
	const activeTab = cur.u16();
	const firstTab = cur.u16();
	cur.skip(2);
	return { activeTab, firstTab, flags, tabRatio: cur.u16() };
}

/** Parse the records of the globals substream (BOF first). */
export function parseGlobals(records: readonly BiffRecord[], warnings: XlsWarnings): XlsGlobals {
	const g: XlsGlobals = {
		date1904: false,
		sheets: [],
		sst: { strings: [], runs: new Map() },
		fonts: [],
		formats: new Map(),
		xfs: [],
		styles: [],
		names: [],
		supbooks: [],
		xti: [],
	};
	const protection: WorkbookProtection = {};
	for (const record of records) {
		try {
			switch (record.type) {
				case RT.DATEMODE:
					g.date1904 = new BiffCursor(record.data).u16() === 1;
					break;
				case RT.BOUNDSHEET8:
					g.sheets.push(parseBoundSheet(record));
					break;
				case RT.SST: {
					const sst = parseSst(record);
					g.sst = sst;
					if (sst.truncated) {
						warnings.warn('truncated-record', 'The shared string table is incomplete');
					}
					break;
				}
				case RT.FONT:
					g.fonts.push(parseFont(record));
					break;
				case RT.FORMAT: {
					const f = parseFormat(record);
					g.formats.set(f.id, f.code);
					break;
				}
				case RT.XF:
					g.xfs.push(parseXf(record));
					break;
				case RT.STYLE:
					g.styles.push(parseStyle(record));
					break;
				case RT.PALETTE:
					g.palette = parsePalette(record);
					break;
				case RT.NAME:
					g.names.push(parseName(record));
					break;
				case RT.SUPBOOK:
					g.supbooks.push(parseSupBook(record));
					break;
				case RT.EXTERNNAME: {
					const cur = new BiffCursor(record.data, record.breaks);
					cur.skip(6);
					g.supbooks[g.supbooks.length - 1]?.names.push(cur.shortString());
					break;
				}
				case RT.EXTERNSHEET:
					g.xti = parseExternSheet(record);
					break;
				case RT.WINDOW1:
					g.window ??= parseWindow1(record);
					break;
				case RT.WRITEACCESS: {
					const name = new BiffCursor(record.data, record.breaks).unicodeString().trim();
					if (name) {
						g.lastModifiedBy = name;
					}
					break;
				}
				case RT.PROTECT:
					if (new BiffCursor(record.data).u16() === 1) {
						protection.lockStructure = true;
					}
					break;
				case RT.WINDOWPROTECT:
					if (new BiffCursor(record.data).u16() === 1) {
						protection.lockWindows = true;
					}
					break;
				case RT.PASSWORD: {
					const hash = new BiffCursor(record.data).u16();
					if (hash) {
						protection.workbookPassword = hash.toString(16).toUpperCase().padStart(4, '0');
					}
					break;
				}
			}
		} catch (e) {
			warnings.corrupt(`record 0x${record.type.toString(16)} of the workbook globals`, e);
		}
	}
	if (Object.keys(protection).length > 0) {
		g.protection = protection;
	}
	return g;
}

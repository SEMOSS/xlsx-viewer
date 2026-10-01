/**
 * Load pipeline for legacy Excel 97-2003 workbooks: an OLE2 compound file
 * whose `Workbook` stream holds BIFF8 records -> the same `Workbook` model
 * `readWorkbook` builds from an `.xlsx`. The result has no source package
 * and is marked dirty, so `writeWorkbook` regenerates every part.
 */
import { readCompoundFile } from '../../package/cfb/read';
import { isCompoundFile } from '../../package/opc-package';
import { createDefaultTheme } from '../../styles/defaults';
import { createWorksheet } from '../../sheet/worksheet';
import type { Worksheet } from '../../types/sheet';
import type {
	CalcProperties,
	DocumentProperties,
	Workbook,
	WorkbookView,
} from '../../types/workbook';
import { XlsxFormatError, XlsxPasswordError } from '../errors';
import { BIFF5_VERSION, BOF_TYPE, OLD_BOF_TYPES, RT } from './biff-opcodes';
import { bofInfo, readSubstream } from './biff-records';
import { readDocumentSummary } from './doc-summary';
import { parseGlobals, type XlsGlobals } from './globals';
import { buildDefinedNames, formulaEnv, nameArea } from './names';
import { readWorksheet, type SheetResult } from './sheet-reader';
import { XlsWarnings } from './warnings';
import { buildStyles } from './xls-styles';

export interface ReadXlsOptions {
	/** Reference date; accepted for parity with the `.xlsx` reader's options. */
	now?: Date;
}

const ENCRYPTED_MESSAGE =
	'This workbook is encrypted with a password. ' +
	'Opening encrypted .xls workbooks is not supported yet.';

interface Container {
	stream: Uint8Array;
	streams?: (path: string) => Uint8Array | undefined;
	/** The file carries a VBA project (`_VBA_PROJECT_CUR` storage). */
	vba?: boolean;
}

function damaged(e: unknown): XlsxFormatError {
	const reason = e instanceof Error ? e.message : String(e);
	return new XlsxFormatError(`The .xls file is damaged and cannot be opened (${reason})`);
}

/** The BIFF stream: `Workbook` (BIFF8) or `Book` (BIFF5) in a compound file, or a bare stream. */
function openContainer(bytes: Uint8Array): Container {
	if (!isCompoundFile(bytes)) {
		const type = bytes.length >= 4 ? bytes[0] | (bytes[1] << 8) : -1;
		if (type === RT.BOF) {
			return { stream: bytes };
		}
		if (OLD_BOF_TYPES.includes(type)) {
			throw new XlsxFormatError('Excel 2.x-4.x worksheets are not supported');
		}
		throw new XlsxFormatError('The file is not an Excel 97-2003 workbook');
	}
	try {
		const cfb = readCompoundFile(bytes);
		const stream = cfb.stream('Workbook') ?? cfb.stream('Book');
		if (stream) {
			const vba = cfb.entries.some((e) => e.path.toLowerCase().startsWith('_vba_project_cur'));
			return { stream, streams: (path) => cfb.stream(path), vba };
		}
		if (cfb.stream('EncryptionInfo') || cfb.stream('EncryptedPackage')) {
			throw new XlsxPasswordError(ENCRYPTED_MESSAGE, 'unsupported');
		}
	} catch (e) {
		if (e instanceof XlsxPasswordError) {
			throw e;
		}
		throw damaged(e);
	}
	throw new XlsxFormatError('The file does not contain an Excel workbook stream');
}

/** Globals substream with the version and encryption checks. */
function readGlobals(stream: Uint8Array, warnings: XlsWarnings): XlsGlobals {
	const globals = readSubstream(stream, 0);
	const bof = globals.records[0];
	if (!bof || bof.type !== RT.BOF) {
		throw new XlsxFormatError('The workbook stream does not start with a BOF record');
	}
	const { version, type } = bofInfo(bof);
	if (version === BIFF5_VERSION || (version > 0 && version < BIFF5_VERSION)) {
		throw new XlsxFormatError('Excel 5.0/95 workbooks are not supported');
	}
	if (type !== BOF_TYPE.globals) {
		throw new XlsxFormatError('The workbook stream has no workbook globals');
	}
	if (globals.records.some((r) => r.type === RT.FILEPASS)) {
		throw new XlsxPasswordError(ENCRYPTED_MESSAGE, 'unsupported');
	}
	if (globals.truncated) {
		warnings.warn('truncated-record', 'The workbook globals are incomplete in the file');
	}
	return parseGlobals(globals.records, warnings);
}

const SKIPPED_KIND: Readonly<Record<number, string>> = {
	1: 'macro sheet',
	2: 'chart sheet',
	6: 'Visual Basic module',
};

function workbookView(g: XlsGlobals, index: Map<number, number>): WorkbookView {
	const w = g.window;
	const view: WorkbookView = { activeTab: index.get(w?.activeTab ?? 0) ?? 0 };
	if (!w) {
		return view;
	}
	const first = index.get(w.firstTab);
	if (first) {
		view.firstSheet = first;
	}
	if (!(w.flags & 0x08)) {
		view.showHorizontalScroll = false;
	}
	if (!(w.flags & 0x10)) {
		view.showVerticalScroll = false;
	}
	if (!(w.flags & 0x20)) {
		view.showSheetTabs = false;
	}
	if (w.tabRatio > 0 && w.tabRatio !== 600) {
		view.tabRatio = w.tabRatio;
	}
	return view;
}

/** Read an `.xls` (BIFF8) workbook. Throws `XlsxFormatError` / `XlsxPasswordError`. */
export function readXlsWorkbook(bytes: Uint8Array, _options: ReadXlsOptions = {}): Workbook {
	const container = openContainer(bytes);
	const stream = container.stream;
	const warnings = new XlsWarnings();
	const g = readGlobals(stream, warnings);
	const styles = buildStyles({
		fonts: g.fonts,
		formats: g.formats,
		xfs: g.xfs,
		styles: g.styles,
		palette: g.palette,
	});
	const env = formulaEnv(g);
	const sheets: Worksheet[] = [];
	const results: SheetResult[] = [];
	const index = new Map<number, number>();
	g.sheets.forEach((entry, i) => {
		const kind =
			entry.type === 0 ? undefined : (SKIPPED_KIND[entry.type] ?? 'sheet of an unknown kind');
		const sub =
			kind || entry.offset >= stream.length ? undefined : readSubstream(stream, entry.offset);
		const bof = sub?.records[0];
		const bofType = bof && bof.type === RT.BOF ? bofInfo(bof).type : undefined;
		if (!sub || bofType !== BOF_TYPE.worksheet) {
			const what =
				kind ?? (bofType === BOF_TYPE.chart ? 'chart sheet' : 'sheet that could not be found');
			warnings.warn(
				'sheet-skipped',
				`The ${what} "${entry.name}" was not imported`,
				{ sheet: entry.name },
				entry.name,
			);
			return;
		}
		const result = readWorksheet(
			sub,
			{ name: entry.name, sheetId: sheets.length + 1, state: entry.state },
			{ styles, sst: g.sst, env, warnings },
		);
		if (result.dialog) {
			warnings.warn(
				'sheet-skipped',
				`The dialog sheet "${entry.name}" was not imported`,
				{ sheet: entry.name },
				entry.name,
			);
			return;
		}
		index.set(i, sheets.length);
		sheets.push(result.sheet);
		results.push(result);
	});
	if (sheets.length === 0) {
		warnings.warn('no-worksheets', 'The workbook has no worksheets that could be imported');
		sheets.push(createWorksheet('Sheet1', 1));
	}
	if (container.vba) {
		warnings.warn('macros-dropped', 'The workbook contains VBA macros, which are not imported');
	}
	if (styles.badXf) {
		warnings.warn(
			'missing-style',
			'Some cells refer to missing formats and use the default format',
		);
	}

	const definedNames = buildDefinedNames(g, env, (i) => index.get(i), warnings);
	for (const name of g.names) {
		const local = name.itab > 0 ? index.get(name.itab - 1) : undefined;
		const result = local === undefined ? undefined : results[local];
		const ref =
			result?.autoFilter && name.name === '_xlnm._FilterDatabase' ? nameArea(name.rgce) : undefined;
		if (result && ref) {
			result.sheet.autoFilter = { ref, columns: [] };
		}
	}

	const calc: CalcProperties = results[0]?.calc ?? {};
	const properties: DocumentProperties = container.streams
		? readDocumentSummary(container.streams)
		: {};
	if (g.lastModifiedBy && !properties.lastModifiedBy) {
		properties.lastModifiedBy = g.lastModifiedBy;
	}
	const workbook: Workbook = {
		sheets,
		definedNames,
		styles: styles.sheet,
		theme: createDefaultTheme(),
		properties,
		view: workbookView(g, index),
		date1904: g.date1904,
		calc,
		media: new Map(),
		persons: [],
		warnings: warnings.list,
		dirty: true,
	};
	if (g.protection) {
		workbook.protection = g.protection;
	}
	return workbook;
}

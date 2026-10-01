import {
	OpcPackage,
	isCompoundFile,
	isZipFile,
	toUint8Array,
	type BinaryInput,
} from '../package/opc-package';
import { decryptPackage } from '../package/crypto';
import { relsPathFor } from '../package/paths';
import { findRelByType } from '../package/relationships';
import type { DocumentProperties, Workbook } from '../types/workbook';
import type { Relationship, Worksheet } from '../types/sheet';
import { CT, REL } from '../xml/namespaces';
import { attempt, ReadContext, type ReadOptions } from './context';
import { readDocumentProperties } from './doc-props';
import { compoundFileError, isEncryptedPackage, XlsxFormatError } from './errors';
import { readXlsWorkbook } from './xls';
import { csvToWorkbook } from '../convert/csv-import';
import { parseDynamicArrayMetadata } from './metadata';
import { readPartText, relLookup, relTypeSuffix, type RelLookup } from './part-io';
import { parseSharedStrings } from './shared-strings';
import { parseStyleSheet } from './styles';
import { parseTheme } from './theme';
import { parseWorkbookPart, sheetKindFromRelType, type SheetEntry } from './workbook-part';
import { createSheetShell, readSheetPart } from './worksheet';
import { parsePersons } from './comments';
import { parseSpreadsheetXml } from './xml-normalize';
import { lname } from './xml-utils';

const WORKBOOK_CONTENT_TYPES: readonly string[] = [
	CT.workbook,
	CT.workbookMacro,
	CT.workbookTemplate,
	'application/vnd.ms-excel.template.macroEnabled.main+xml',
	'application/vnd.ms-excel.addin.macroEnabled.main+xml',
];

/** Relationships of a part; a malformed `.rels` part is recorded and treated as empty. */
async function safeRels(pkg: OpcPackage, part: string, ctx?: ReadContext): Promise<Relationship[]> {
	try {
		return await pkg.rels(part);
	} catch (e) {
		ctx?.unreadable(part === '' ? '_rels/.rels' : relsPathFor(part), e, 'relationships');
		return [];
	}
}

async function findWorkbookPart(pkg: OpcPackage): Promise<string | undefined> {
	const main = findRelByType(await safeRels(pkg, ''), REL.officeDocument);
	if (main && main.targetMode !== 'External') {
		const path = pkg.resolve('', main);
		if (pkg.has(path)) {
			return path;
		}
	}
	for (const type of WORKBOOK_CONTENT_TYPES) {
		const parts = pkg.contentTypes.partsOfType(type).filter((p) => pkg.has(p));
		if (parts.length > 0) {
			return parts[0];
		}
	}
	return pkg.has('xl/workbook.xml') ? 'xl/workbook.xml' : undefined;
}

async function openPackage(
	data: BinaryInput,
	password?: string,
): Promise<{ pkg: OpcPackage; encrypted: boolean }> {
	let bytes = toUint8Array(data);
	let encrypted = false;
	if (isCompoundFile(bytes)) {
		if (!isEncryptedPackage(bytes)) {
			throw compoundFileError(bytes);
		}
		// Encrypt with Password: the ZIP package is inside an encrypted compound file.
		bytes = await decryptPackage(bytes, password);
		encrypted = true;
	}
	if (!isZipFile(bytes)) {
		throw new XlsxFormatError('The file is not an Excel workbook (.xlsx files are ZIP packages)');
	}
	try {
		return { pkg: await OpcPackage.open(bytes), encrypted };
	} catch (e) {
		const reason = e instanceof Error ? e.message : String(e);
		throw new XlsxFormatError(`The workbook file is damaged and cannot be opened (${reason})`);
	}
}

/** Load a workbook-level part through the workbook relationships, recording failures as warnings. */
async function loadRelated<T>(
	ctx: ReadContext,
	rels: RelLookup,
	type: string,
	what: string,
	parse: (xml: string) => T,
): Promise<T | undefined> {
	const rel = findRelByType(rels.rels, type);
	const path = rel ? rels.path(rel.id) : undefined;
	if (!path) {
		return undefined;
	}
	return attempt(
		ctx,
		path,
		what,
		async () => {
			const xml = await readPartText(ctx.pkg, path);
			if (xml === undefined) {
				ctx.warn('missing-part', `The ${what} part ${path} is missing`, { part: path });
				return undefined;
			}
			return parse(xml);
		},
		undefined,
	);
}

async function readSheet(ctx: ReadContext, entry: SheetEntry, rels: RelLookup): Promise<Worksheet> {
	const rel = rels.byId(entry.relId);
	const sheet = createSheetShell(entry, sheetKindFromRelType(rel?.type));
	const partPath = rels.path(entry.relId);
	if (!partPath) {
		ctx.warn('missing-part', `Sheet "${entry.name}" has no part in the package`, {
			sheet: entry.name,
		});
		return sheet;
	}
	try {
		await readSheetPart(ctx, sheet, partPath);
	} catch (e) {
		ctx.unreadable(partPath, e, 'sheet', entry.name);
	}
	return sheet;
}

function workbookWarnings(ctx: ReadContext, rels: RelLookup, externalReferences: boolean): void {
	const types = new Set(rels.rels.map((r) => relTypeSuffix(r.type)));
	if (externalReferences || types.has('externalLink')) {
		ctx.warn(
			'external-link',
			'The workbook links to other workbooks; linked values are not refreshed',
		);
	}
	if (types.has('slicerCache') || types.has('timelineCache')) {
		ctx.warn('slicer', 'The workbook contains slicers or timelines, which are not interactive');
	}
	if (types.has('connections')) {
		ctx.warn('external-data', 'The workbook has external data connections; data is not refreshed');
	}
}

/** Read `.xlsx` / `.xlsm` / `.xltx` bytes into a {@link Workbook}. */
export async function readWorkbook(
	data: BinaryInput,
	options: ReadOptions = {},
): Promise<Workbook> {
	const raw = toUint8Array(data);
	const textFile = /\.(csv|tsv|txt)$/i.exec(options.fileName ?? '');
	if (textFile && !isZipFile(raw) && !isCompoundFile(raw)) {
		const sheetName = (options.fileName ?? '').replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '');
		const text = new TextDecoder().decode(raw);
		return csvToWorkbook(text, {
			sheetName,
			delimiter: textFile[1].toLowerCase() === 'tsv' ? '\t' : undefined,
		});
	}
	if (isCompoundFile(raw) && !isEncryptedPackage(raw)) {
		// A compound file that is not an encrypted package: a legacy .xls (BIFF8) workbook.
		return readXlsWorkbook(raw, { now: options.now });
	}
	const { pkg, encrypted } = await openPackage(raw, options.password);
	const workbookPath = await findWorkbookPart(pkg);
	if (!workbookPath) {
		const mimetype = await pkg.text('mimetype');
		if (mimetype?.includes('opendocument')) {
			throw new XlsxFormatError('OpenDocument spreadsheets (.ods) are not supported');
		}
		throw new XlsxFormatError('The package does not contain a workbook (is it an Excel file?)');
	}
	if (workbookPath.toLowerCase().endsWith('.bin')) {
		throw new XlsxFormatError('Excel binary workbooks (.xlsb) are not supported');
	}
	const xml = await readPartText(pkg, workbookPath);
	let root;
	try {
		root = parseSpreadsheetXml(xml ?? '').root;
	} catch (e) {
		const reason = e instanceof Error ? e.message : String(e);
		throw new XlsxFormatError(`The workbook part is not valid XML (${reason})`);
	}
	if (lname(root) !== 'workbook') {
		throw new XlsxFormatError(`The main part is <${root.name}>, not a SpreadsheetML workbook`);
	}

	const ctx = new ReadContext(pkg, options);
	const wb = parseWorkbookPart(root);
	ctx.date1904 = wb.date1904;
	const rels = relLookup(workbookPath, await safeRels(pkg, workbookPath, ctx), pkg);

	const [theme, styles, sharedStrings, dynamicArrayCm, persons] = await Promise.all([
		loadRelated(ctx, rels, REL.theme, 'theme', parseTheme),
		loadRelated(ctx, rels, REL.styles, 'styles', parseStyleSheet),
		loadRelated(ctx, rels, REL.sharedStrings, 'shared strings', parseSharedStrings),
		loadRelated(ctx, rels, REL.sheetMetadata, 'metadata', parseDynamicArrayMetadata),
		loadRelated(ctx, rels, REL.person, 'persons', parsePersons),
	]);
	if (theme) {
		ctx.theme = theme;
	}
	if (styles) {
		ctx.styles = styles;
	}
	if (sharedStrings) {
		ctx.sharedStrings = sharedStrings;
	}
	ctx.dynamicArrayCm = dynamicArrayCm;

	const sheets: Worksheet[] = [];
	for (const entry of wb.sheets) {
		sheets.push(await readSheet(ctx, entry, rels));
	}
	workbookWarnings(ctx, rels, wb.externalReferences);
	const properties: DocumentProperties = await readDocumentProperties(ctx);

	const workbook: Workbook = {
		sheets,
		definedNames: wb.definedNames,
		styles: ctx.styles,
		theme: ctx.theme,
		properties,
		view: wb.view,
		date1904: wb.date1904,
		calc: wb.calc,
		media: ctx.media,
		persons: persons ?? [],
		warnings: ctx.warnings,
		hasMacros:
			rels.rels.some((r) => relTypeSuffix(r.type) === 'vbaProject') || pkg.has('xl/vbaProject.bin'),
		sourcePackage: pkg,
		dirty: false,
	};
	if (wb.protection) {
		workbook.protection = wb.protection;
	}
	if (encrypted) {
		workbook.encrypted = true;
	}
	return workbook;
}

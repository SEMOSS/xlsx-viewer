import JSZip from 'jszip';

/**
 * Test helpers: build `.xlsx` packages in memory from XML strings.
 *
 * ```ts
 * const bytes = await buildWorkbook({ sheets: [{ name: 'Data', xml: worksheet('<row r="1">...</row>') }] });
 * const wb = await readWorkbook(bytes);
 * ```
 */

export const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

export interface RelSpec {
	id: string;
	/** Full type URI, or a short suffix appended to the transitional base. */
	type: string;
	target: string;
	external?: boolean;
}

export interface SheetSpec {
	name: string;
	/** Complete sheet part XML. */
	xml: string;
	rels?: RelSpec[];
	kind?: 'worksheet' | 'chartsheet' | 'dialogsheet';
	state?: 'hidden' | 'veryHidden';
}

export interface WorkbookSpec {
	sheets: SheetSpec[];
	/** Plain shared strings, or a complete `sst` part. */
	sharedStrings?: string[] | string;
	styles?: string;
	theme?: string;
	/** Markup inserted into workbook.xml after `<sheets>` (definedNames, calcPr, ...). */
	workbookAfterSheets?: string;
	/** Markup inserted into workbook.xml before `<sheets>` (workbookPr, bookViews, ...). */
	workbookBeforeSheets?: string;
	/** Extra workbook relationships. */
	workbookRels?: RelSpec[];
	/** Extra parts (path -> content). */
	parts?: Record<string, string | Uint8Array>;
	/** Extra package relationships (`_rels/.rels`). */
	packageRels?: RelSpec[];
}

export function relsXml(rels: RelSpec[]): string {
	const items = rels
		.map((r) => {
			const type = r.type.includes('://') ? r.type : `${REL_BASE}/${r.type}`;
			const mode = r.external ? ' TargetMode="External"' : '';
			return `<Relationship Id="${r.id}" Type="${type}" Target="${escapeAttr(r.target)}"${mode}/>`;
		})
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

function escapeAttr(s: string): string {
	return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

export function escapeText(s: string): string {
	return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A worksheet part with the given `<sheetData>` content and surrounding markup. */
export function worksheet(
	sheetData: string,
	opts: { before?: string; after?: string } = {},
): string {
	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
		`<worksheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}" xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}" ` +
		`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
		`${opts.before ?? ''}<sheetData>${sheetData}</sheetData>${opts.after ?? ''}</worksheet>`
	);
}

export function sharedStringsXml(strings: string[]): string {
	const items = strings
		.map((s) => `<si><t xml:space="preserve">${escapeText(s)}</t></si>`)
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${MAIN_NS}" count="${strings.length}" uniqueCount="${strings.length}">${items}</sst>`;
}

const SHEET_REL: Record<string, string> = {
	worksheet: 'worksheet',
	chartsheet: 'chartsheet',
	dialogsheet: 'dialogsheet',
};

/** All parts of a workbook package described by `spec`. */
export function workbookParts(spec: WorkbookSpec): Record<string, string | Uint8Array> {
	const parts: Record<string, string | Uint8Array> = {};
	const wbRels: RelSpec[] = [];
	const sheetEls: string[] = [];
	spec.sheets.forEach((s, i) => {
		const kind = s.kind ?? 'worksheet';
		const folder =
			kind === 'chartsheet'
				? 'chartsheets'
				: kind === 'dialogsheet'
					? 'dialogsheets'
					: 'worksheets';
		const path = `xl/${folder}/sheet${i + 1}.xml`;
		parts[path] = s.xml;
		if (s.rels?.length) {
			parts[`xl/${folder}/_rels/sheet${i + 1}.xml.rels`] = relsXml(s.rels);
		}
		const rid = `rIdS${i + 1}`;
		wbRels.push({ id: rid, type: SHEET_REL[kind], target: `${folder}/sheet${i + 1}.xml` });
		const state = s.state ? ` state="${s.state}"` : '';
		sheetEls.push(`<sheet name="${escapeAttr(s.name)}" sheetId="${i + 1}"${state} r:id="${rid}"/>`);
	});
	if (spec.sharedStrings !== undefined) {
		parts['xl/sharedStrings.xml'] = Array.isArray(spec.sharedStrings)
			? sharedStringsXml(spec.sharedStrings)
			: spec.sharedStrings;
		wbRels.push({ id: 'rIdSST', type: 'sharedStrings', target: 'sharedStrings.xml' });
	}
	if (spec.styles !== undefined) {
		parts['xl/styles.xml'] = spec.styles;
		wbRels.push({ id: 'rIdStyles', type: 'styles', target: 'styles.xml' });
	}
	if (spec.theme !== undefined) {
		parts['xl/theme/theme1.xml'] = spec.theme;
		wbRels.push({ id: 'rIdTheme', type: 'theme', target: 'theme/theme1.xml' });
	}
	wbRels.push(...(spec.workbookRels ?? []));
	parts['xl/workbook.xml'] =
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${MAIN_NS}" xmlns:r="${R_NS}">` +
		`${spec.workbookBeforeSheets ?? ''}<sheets>${sheetEls.join('')}</sheets>${spec.workbookAfterSheets ?? ''}</workbook>`;
	parts['xl/_rels/workbook.xml.rels'] = relsXml(wbRels);
	parts['_rels/.rels'] = relsXml([
		{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' },
		...(spec.packageRels ?? []),
	]);
	parts['[Content_Types].xml'] =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		'<Default Extension="png" ContentType="image/png"/>' +
		'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
		'</Types>';
	Object.assign(parts, spec.parts ?? {});
	return parts;
}

/** Zip a map of part path -> content. */
export async function zipParts(parts: Record<string, string | Uint8Array>): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [path, content] of Object.entries(parts)) {
		zip.file(path, content);
	}
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** Build and zip a workbook package. */
export async function buildWorkbook(spec: WorkbookSpec): Promise<Uint8Array> {
	return zipParts(workbookParts(spec));
}

/** A 1x1 transparent PNG. */
export const PNG_1X1 = new Uint8Array([
	0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
	0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
	0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
	0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
	0x42, 0x60, 0x82,
]);

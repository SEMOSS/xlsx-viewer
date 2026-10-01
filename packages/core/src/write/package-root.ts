/** Package-level parts: `_rels/.rels` and the document property parts. */
import { findRelByType, resolveRelTarget } from '../package';
import { localName, parseXml, textContent } from '../xml';
import { CT, REL } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { addRel, type WriteContext } from './context';
import { buildAppXml, buildCoreXml, buildCustomXml, customChanged } from './doc-props';
import type { WorkbookPart } from './workbook-part';
import type { WorkbookSheetEntry } from './workbook-xml';

async function parsed(ctx: WriteContext, path: string): Promise<XmlElement | undefined> {
	const xml = await ctx.sourceText(path);
	if (!xml) {
		return undefined;
	}
	try {
		return parseXml(xml).root;
	} catch {
		return undefined;
	}
}

/** Company or Manager differ from the original `docProps/app.xml`. */
function appPropsChanged(
	original: XmlElement | undefined,
	company?: string,
	manager?: string,
): boolean {
	if (!original) {
		return false;
	}
	const text = (name: string): string => {
		const child = original.children.find(
			(c) => typeof c === 'object' && 'name' in c && localName(c.name) === name,
		);
		return child && typeof child === 'object' && 'name' in child ? textContent(child) : '';
	};
	return (
		(company !== undefined && company !== text('Company')) ||
		(manager !== undefined && manager !== text('Manager'))
	);
}

/**
 * Write the document properties and the package relationships. `changed`
 * tells whether any other part changed (core.xml then gets a new modified time).
 */
export async function writePackageRoot(
	ctx: WriteContext,
	wb: WorkbookPart,
	entries: readonly WorkbookSheetEntry[],
	changed: boolean,
): Promise<void> {
	const rels = (await ctx.rels('')).map((r) => ({ ...r }));
	if (!findRelByType(rels, REL.officeDocument)) {
		addRel(rels, REL.officeDocument, wb.path);
	}
	const fresh = !ctx.source || !!ctx.options.regenerateAll;
	const props = ctx.workbook.properties;

	const coreRel = findRelByType(rels, REL.coreProperties);
	const corePath = coreRel ? resolveRelTarget('', coreRel.target) : 'docProps/core.xml';
	if (fresh || changed) {
		const original = coreRel ? await parsed(ctx, corePath) : undefined;
		ctx.writeText(corePath, buildCoreXml(props, original, ctx.now), CT.coreProperties);
		if (!coreRel) {
			addRel(rels, REL.coreProperties, corePath);
		}
	}

	const appRel = findRelByType(rels, REL.extendedProperties);
	const appPath = appRel ? resolveRelTarget('', appRel.target) : 'docProps/app.xml';
	const names = entries.filter((e) => e.sheet.kind === 'worksheet').map((e) => e.name);
	const previous = wb.originalSheets.filter((s) => s.worksheet).map((s) => s.name);
	const namesChanged = names.join('\u0000') !== previous.join('\u0000');
	const original = appRel && !ctx.options.regenerateAll ? await parsed(ctx, appPath) : undefined;
	if (
		fresh ||
		namesChanged ||
		appPropsChanged(original, props.company, props.manager) ||
		(!appRel && changed)
	) {
		const xml = buildAppXml(
			{
				sheetNames: names,
				previousSheetNames: previous,
				application: ctx.options.application ?? props.application ?? 'Microsoft Excel',
				appVersion: props.appVersion,
				company: props.company,
				manager: props.manager,
			},
			original,
		);
		ctx.writeText(appPath, xml, CT.extendedProperties);
		if (!appRel) {
			addRel(rels, REL.extendedProperties, appPath);
		}
	}

	const custom = props.custom ?? [];
	const customRel = findRelByType(rels, REL.customProperties);
	if (custom.length) {
		const customPath = customRel ? resolveRelTarget('', customRel.target) : 'docProps/custom.xml';
		const originalXml = customRel ? await ctx.sourceText(customPath) : undefined;
		if (!originalXml || customChanged(custom, originalXml)) {
			ctx.writeText(customPath, buildCustomXml(custom), CT.customProperties);
		}
		if (!customRel) {
			addRel(rels, REL.customProperties, customPath);
		}
	}
	await ctx.writeRelsIfChanged('', rels);
}

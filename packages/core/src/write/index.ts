/**
 * Save pipeline: `Workbook` -> `.xlsx` bytes.
 *
 * The output starts as a copy of the source package (when the workbook was
 * loaded), so parts the engine does not understand are carried over. Parts
 * are regenerated according to the model's change tracking:
 *
 * - sheets with `dirty === false` keep their part and relationships;
 * - drawings / comments / tables are rewritten when their flags say so;
 * - styles and shared strings are append-only (original entries verbatim);
 * - `workbook.xml` is rebuilt from the original root and only written when
 *   it differs; the calc chain is dropped when cells were rewritten;
 * - parts orphaned by the save (deleted sheets and what they linked) are removed.
 *
 * With `options.password` the finished ZIP is encrypted as Excel's Encrypt with
 * Password does (ECMA-376 Agile encryption inside a compound file).
 */
import { encryptPackage } from '../package/crypto';
import type { Workbook } from '../types/workbook';
import { asOpcPackage, WriteContext, type WriteOptions } from './context';
import { removeOrphans } from './package-gc';
import { writePackageRoot } from './package-root';
import { createWriterState, writeSharedParts } from './shared-parts';
import { type SheetTarget, writeWorksheet } from './sheet-writer';
import {
	assignSheetParts,
	loadWorkbookPart,
	removeCalcChain,
	writeWorkbookPart,
} from './workbook-part';

export type { WriteOptions } from './context';

function hasStale(target: SheetTarget): boolean {
	for (const c of target.sheet.cells.values()) {
		if (c.stale) {
			return true;
		}
	}
	return false;
}

/** Serialize a workbook to `.xlsx` bytes (encrypted when `options.password` is set). */
export async function writeWorkbook(
	workbook: Workbook,
	options: WriteOptions = {},
): Promise<Uint8Array> {
	const source = asOpcPackage(workbook.sourcePackage);
	const ctx = new WriteContext(workbook, source, options);
	const wb = await loadWorkbookPart(ctx);
	const targets = assignSheetParts(ctx, wb);
	const state = await createWriterState(ctx, wb);

	const regenerated = new Set<SheetTarget>();
	let needsCalc = false;
	for (const target of targets) {
		if (target.hasSource && target.sheet.kind !== 'worksheet') {
			// Chart sheets and legacy sheet kinds are carried over as they are.
			if (target.sheet.dirty !== false) {
				ctx.warn(
					'sheet-kept',
					`Changes to ${target.sheet.kind} "${target.sheet.name}" are not saved`,
					{
						sheet: target.sheet.name,
					},
				);
			}
			continue;
		}
		const result = await writeWorksheet(state, target);
		if (result.regenerated) {
			regenerated.add(target);
		}
		needsCalc ||= result.needsCalc;
	}
	if (!needsCalc) {
		needsCalc = targets.some((t) => !regenerated.has(t) && hasStale(t));
	}

	await writeSharedParts(ctx, wb, state, targets, regenerated);
	// The calc chain lists formula cells by sheet id: stale once cells were rewritten or a sheet went away.
	const kept = new Set(targets.map((t) => t.partPath.toLowerCase()));
	const sheetRemoved = wb.originalSheets.some(
		(s) => s.partPath && !kept.has(s.partPath.toLowerCase()),
	);
	if (ctx.source && (regenerated.size || sheetRemoved)) {
		removeCalcChain(ctx, wb);
	}
	const entries = await writeWorkbookPart(ctx, wb, targets, needsCalc);
	await writePackageRoot(ctx, wb, entries, ctx.changed);
	await removeOrphans(ctx);
	const zip = await ctx.writer.generate();
	return options.password ? encryptPackage(zip, options.password) : zip;
}

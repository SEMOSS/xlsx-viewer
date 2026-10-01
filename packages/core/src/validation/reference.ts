import { parseRange } from '../address/range';
import { findSheetIndex } from '../sheet/workbook';
import type { RangeAddress } from '../types/address';
import type { DefinedName, Workbook } from '../types/workbook';

/** A plain cell / range reference resolved to a sheet. */
export interface ResolvedReference {
	sheetIndex: number;
	range: RangeAddress;
	/** The reference came from a defined name. */
	named: boolean;
}

const REFERENCE =
	/^(?:(?:'((?:[^']|'')+)'|([^'!:\s()"=<>&,;+*/^-]+))!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})$/;

const NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

/** A defined name visible from a sheet: sheet-scoped names win over workbook-scoped ones. */
export function findDefinedName(
	workbook: Workbook,
	name: string,
	sheetIndex: number,
): DefinedName | undefined {
	const lower = name.toLowerCase();
	let global: DefinedName | undefined;
	for (const dn of workbook.definedNames) {
		if (dn.name.toLowerCase() !== lower) {
			continue;
		}
		if (dn.localSheet === sheetIndex) {
			return dn;
		}
		if (dn.localSheet === undefined) {
			global ??= dn;
		}
	}
	return global;
}

/** Whether text is the name of a defined name visible from the sheet. */
export function isDefinedNameReference(
	text: string,
	workbook: Workbook,
	sheetIndex: number,
): boolean {
	const t = text.trim();
	return NAME.test(t) && !!findDefinedName(workbook, t, sheetIndex);
}

/**
 * Resolve `A1:B5`, `$A:$A`, `Sheet2!A1:A9`, `'My Sheet'!$B$2` or a defined
 * name whose formula is such a reference. Returns `undefined` for anything
 * else (functions, unions, unknown sheets), which callers hand to the
 * formula evaluator instead.
 */
export function resolveReference(
	text: string,
	workbook: Workbook,
	sheetIndex: number,
	depth = 0,
): ResolvedReference | undefined {
	const t = text.trim().replace(/^=/, '');
	const m = REFERENCE.exec(t);
	if (m) {
		const sheetName = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
		const index = sheetName === undefined ? sheetIndex : findSheetIndex(workbook, sheetName);
		const range = parseRange(m[3]);
		return index >= 0 && range ? { sheetIndex: index, range, named: false } : undefined;
	}
	if (depth < 4 && NAME.test(t)) {
		const dn = findDefinedName(workbook, t, sheetIndex);
		if (dn) {
			const inner = resolveReference(dn.formula, workbook, dn.localSheet ?? sheetIndex, depth + 1);
			return inner ? { ...inner, named: true } : undefined;
		}
	}
	return undefined;
}

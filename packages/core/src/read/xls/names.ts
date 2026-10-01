/**
 * Defined names (NAME / Lbl records) -> `workbook.definedNames`. Names of
 * macros and user-defined functions (fFunc / fProc) and the `_xlfn.`
 * placeholders Excel writes for newer functions are only formula targets,
 * so they are not listed. Sheet indices are remapped because chart, macro
 * and dialog sheets are not imported.
 */
import type { RangeAddress } from '../../types/address';
import type { DefinedName } from '../../types/workbook';
import {
	decodeRgce,
	type FormulaEnv,
	type Ptg,
	renderFormula,
	XlsExternalReferenceError,
} from './formula';
import type { XlsGlobals } from './globals';
import type { XlsWarnings } from './warnings';

/** Formula lookups from the globals. */
export function formulaEnv(g: XlsGlobals): FormulaEnv {
	return {
		sheetNames: g.sheets.map((s) => s.name),
		xti: g.xti,
		supbooks: g.supbooks,
		names: g.names.map((n) => n.name),
	};
}

/**
 * Build the defined names. `sheetIndex` maps a BOUNDSHEET index to the
 * index of the imported sheet (`undefined` when the sheet was skipped).
 */
export function buildDefinedNames(
	g: XlsGlobals,
	env: FormulaEnv,
	sheetIndex: (biffIndex: number) => number | undefined,
	warnings: XlsWarnings,
): DefinedName[] {
	const out: DefinedName[] = [];
	for (const n of g.names) {
		if (n.func || n.rgce.length === 0 || n.name.toLowerCase().startsWith('_xlfn.')) {
			continue;
		}
		if (!n.name || /[\u0000-\u001f]/.test(n.name)) {
			warnings.warn(
				'name-skipped',
				'A defined name with an invalid name was dropped',
				{},
				'invalid',
			);
			continue;
		}
		const local = n.itab > 0 ? sheetIndex(n.itab - 1) : undefined;
		if (n.itab > 0 && local === undefined) {
			warnings.warn(
				'name-skipped',
				`The name "${n.name}" belongs to a sheet that was not imported and was dropped`,
				{},
				n.name,
			);
			continue;
		}
		let formula: string;
		try {
			formula = renderFormula(decodeRgce(n.rgce, n.extra), env, { row: 0, col: 0, mode: 'name' });
		} catch (e) {
			if (e instanceof XlsExternalReferenceError) {
				warnings.warn(
					'external-link',
					`The name "${n.name}" refers to ${e.source}; ` +
						'external links are not imported, so it was dropped',
					{},
					n.name,
				);
			} else {
				const reason = e instanceof Error ? e.message : String(e);
				warnings.warn(
					'name-skipped',
					`The name "${n.name}" could not be read (${reason})`,
					{},
					n.name,
				);
			}
			continue;
		}
		const dn: DefinedName = { name: n.name, formula };
		if (local !== undefined) {
			dn.localSheet = local;
		}
		if (n.hidden) {
			dn.hidden = true;
		}
		if (n.builtin) {
			dn.builtin = true;
		}
		out.push(dn);
	}
	return out;
}

/**
 * The single area a name refers to (as in `_FilterDatabase`), or
 * `undefined` when its formula is anything else.
 */
export function nameArea(rgce: Uint8Array): RangeAddress | undefined {
	let tokens: Ptg[];
	try {
		tokens = decodeRgce(rgce);
	} catch {
		return undefined;
	}
	const t = tokens.length === 1 ? tokens[0] : undefined;
	if (!t || (t.k !== 'area' && t.k !== 'ref')) {
		return undefined;
	}
	const [a, b] = t.k === 'area' ? [t.first, t.last] : [t.loc, t.loc];
	const col = (v: number): number => v & 0x3fff;
	return {
		top: Math.min(a.rw, b.rw),
		left: Math.min(col(a.col), col(b.col)),
		bottom: Math.max(a.rw, b.rw),
		right: Math.max(col(a.col), col(b.col)),
	};
}

import type { DefinedName, Workbook } from '../types/workbook';
import { rewriteNameReferences } from './names-rename';
import {
	BUILTIN_NAME_PREFIX,
	definedNameDisplayName,
	isValidDefinedNameSyntax,
} from './names-syntax';
import type { EditTransaction } from './transaction';

/**
 * Defined names (Formulas > Define Name / Name Manager): validation and
 * undoable add / edit / delete. Renaming rewrites the formulas that use
 * the name; deleting leaves them as written (they evaluate to `#NAME?`).
 */

export {
	BUILTIN_NAME_PREFIX,
	definedNameDisplayName,
	isValidDefinedNameSyntax,
	looksLikeCellReference,
	sanitizeDefinedName,
} from './names-syntax';
export { renameDefinedNameInFormula, type NameUseMatcher } from './names-rename';
export {
	definedNameDisplayValue,
	selectionReference,
	suggestDefinedName,
	type DefinedNameEvaluator,
} from './names-display';
export {
	parseRefersTo,
	qualifyFormulaReferences,
	REFERS_TO_ERROR_MESSAGE,
} from './names-refers-to';

/** Excel's message for a name that breaks the naming rules. */
export const DEFINED_NAME_INVALID_MESSAGE =
	'The name that you entered is not valid.\n\n' +
	'Reasons for this can include:\n' +
	' - The name does not begin with a letter or an underscore\n' +
	' - The name contains a space or other invalid characters\n' +
	' - The name conflicts with an Excel built-in name or the name of another object in the workbook';

/** Excel's message for a name already defined in the same scope. */
export const DEFINED_NAME_DUPLICATE_MESSAGE = 'The name already exists. Enter a unique name.';

/** Built-in names a user may type; with a sheet scope they are stored as `_xlnm.` names. */
const USER_BUILTINS = ['Print_Area', 'Print_Titles'];

function scopeKey(scope: number | undefined): number {
	return scope ?? -1;
}

/** Same name (built-ins with or without `_xlnm.`) in the same scope. */
function sameDefinition(a: DefinedName, b: DefinedName): boolean {
	return (
		definedNameDisplayName(a).toLowerCase() === definedNameDisplayName(b).toLowerCase() &&
		scopeKey(a.localSheet) === scopeKey(b.localSheet)
	);
}

/**
 * Check a name typed in the New Name / Edit Name dialog for `scope` (a
 * sheet index, or undefined for the workbook). Returns Excel's message, or
 * null when the name can be used. `except` is the definition being edited.
 */
export function validateDefinedName(
	workbook: Workbook,
	name: string,
	scope: number | undefined,
	except?: DefinedName,
): string | null {
	if (!isValidDefinedNameSyntax(name)) {
		return DEFINED_NAME_INVALID_MESSAGE;
	}
	const lower = name.toLowerCase();
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.name.toLowerCase() === lower || table.displayName.toLowerCase() === lower) {
				return DEFINED_NAME_INVALID_MESSAGE;
			}
		}
	}
	for (const def of workbook.definedNames) {
		if (def === except || (except && sameDefinition(def, except))) {
			continue;
		}
		if (scopeKey(def.localSheet) !== scopeKey(scope)) {
			continue;
		}
		if (def.name.toLowerCase() === lower || definedNameDisplayName(def).toLowerCase() === lower) {
			return DEFINED_NAME_DUPLICATE_MESSAGE;
		}
	}
	return null;
}

/** Name Manager order: by name (case-insensitive), then workbook scope before sheet scopes. */
export function compareDefinedNames(a: DefinedName, b: DefinedName): number {
	const na = definedNameDisplayName(a).toLowerCase();
	const nb = definedNameDisplayName(b).toLowerCase();
	if (na !== nb) {
		return na < nb ? -1 : 1;
	}
	return scopeKey(a.localSheet) - scopeKey(b.localSheet);
}

/** Index of the definition with the same name and scope as `target`, or -1. */
export function findDefinedNameIndex(workbook: Workbook, target: DefinedName): number {
	const exact = workbook.definedNames.indexOf(target);
	return exact >= 0 ? exact : workbook.definedNames.findIndex((n) => sameDefinition(n, target));
}

/** Insert at the sorted position without reordering existing entries. */
function insertSorted(list: DefinedName[], def: DefinedName): void {
	const at = list.findIndex((n) => compareDefinedNames(n, def) > 0);
	list.splice(at < 0 ? list.length : at, 0, def);
}

/** Name to store: user-typed built-ins on a sheet become `_xlnm.` names; edits keep the prefix. */
function storedName(def: DefinedName, current?: DefinedName): string {
	const lower = def.name.toLowerCase();
	if (lower.startsWith(BUILTIN_NAME_PREFIX)) {
		return def.name;
	}
	if (current?.builtin && definedNameDisplayName(current).toLowerCase() === lower) {
		return current.name;
	}
	const builtin = USER_BUILTINS.find((b) => b.toLowerCase() === lower);
	return builtin && def.localSheet !== undefined ? BUILTIN_NAME_PREFIX + builtin : def.name;
}

function buildDefinition(def: DefinedName, base?: DefinedName): DefinedName {
	const name = storedName(def, base);
	const out: DefinedName = { name, formula: def.formula.trim().replace(/^=/, '') };
	if (def.localSheet !== undefined) {
		out.localSheet = def.localSheet;
	}
	const hidden = def.hidden ?? base?.hidden;
	if (hidden) {
		out.hidden = true;
	}
	const comment = def.comment?.trim();
	if (comment) {
		out.comment = def.comment;
	}
	if (name.toLowerCase().startsWith(BUILTIN_NAME_PREFIX)) {
		out.builtin = true;
	}
	return out;
}

/**
 * Formulas > Define Name: add a name (formula without the leading `=`).
 * Throws an Error carrying Excel's message when the name is invalid or
 * already defined in its scope. Returns the stored definition.
 */
export function addDefinedName(
	tx: EditTransaction,
	workbook: Workbook,
	def: DefinedName,
): DefinedName {
	const stored = buildDefinition(def);
	const error = validateDefinedName(workbook, definedNameDisplayName(stored), stored.localSheet);
	if (error) {
		throw new Error(error);
	}
	tx.touchWorkbook();
	insertSorted(workbook.definedNames, stored);
	workbook.dirty = true;
	return stored;
}

/**
 * Name Manager > Edit: replace the definition identified by `target`
 * (name + scope) with `next`. A new name is validated (an unchanged one is
 * not) and every formula using the old name is rewritten. Returns the
 * stored definition, or undefined when `target` does not exist.
 */
export function updateDefinedName(
	tx: EditTransaction,
	workbook: Workbook,
	target: DefinedName,
	next: DefinedName,
): DefinedName | undefined {
	const index = findDefinedNameIndex(workbook, target);
	if (index < 0) {
		return undefined;
	}
	const current = workbook.definedNames[index];
	const stored = buildDefinition(next, current);
	const renamed = stored.name !== current.name;
	const moved = scopeKey(stored.localSheet) !== scopeKey(current.localSheet);
	if (stored.name.toLowerCase() !== current.name.toLowerCase() || moved) {
		const error = validateDefinedName(
			workbook,
			definedNameDisplayName(stored),
			stored.localSheet,
			current,
		);
		if (error) {
			throw new Error(error);
		}
	}
	tx.touchWorkbook();
	if (renamed) {
		rewriteNameReferences(tx, workbook, current, stored.name);
	}
	workbook.definedNames.splice(workbook.definedNames.indexOf(current), 1);
	insertSorted(workbook.definedNames, stored);
	workbook.dirty = true;
	return stored;
}

/**
 * Name Manager > Delete. Formulas keep the name text (they recalculate to
 * `#NAME?`). Returns false when `target` does not exist.
 */
export function deleteDefinedName(
	tx: EditTransaction,
	workbook: Workbook,
	target: DefinedName,
): boolean {
	const index = findDefinedNameIndex(workbook, target);
	if (index < 0) {
		return false;
	}
	tx.touchWorkbook();
	workbook.definedNames.splice(index, 1);
	workbook.dirty = true;
	return true;
}

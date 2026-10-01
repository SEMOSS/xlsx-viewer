import type { DefinedName, EditTransaction } from 'xlsx-viewer-core';
import {
	addDefinedName,
	compareDefinedNames,
	definedNameDisplayName,
	definedNameDisplayValue,
	deleteDefinedName,
	selectionReference,
	suggestDefinedName,
	updateDefinedName,
} from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

/**
 * Defined names for the New Name / Edit Name / Name Manager dialogs:
 * dialog defaults, the Name Manager rows and the undoable commands.
 */

/** The fields of the New Name / Edit Name dialog. */
export interface NameDraft {
	name: string;
	/** Sheet index, or undefined for Workbook. */
	scope: number | undefined;
	comment: string;
	/** Refers to text, with its leading `=`. */
	refersTo: string;
}

/** One row of the Name Manager list. */
export interface NameRow {
	key: string;
	def: DefinedName;
	name: string;
	value: string;
	refersTo: string;
	scope: string;
	comment: string;
}

/** Result of a name command: the stored definition, or Excel's message when rejected. */
export interface NameCommandResult {
	def?: DefinedName;
	error?: string;
}

/** Identity of a definition (name and scope), stable across undo. */
export function definedNameKey(def: Pick<DefinedName, 'name' | 'localSheet'>): string {
	return `${def.name.toLowerCase()}|${def.localSheet ?? ''}`;
}

/** New Name defaults: the suggested name, Workbook scope and the selection as Refers to. */
export function newNameDraft(store: ViewerStore): NameDraft {
	const sel = store.selectionOf();
	return {
		name: suggestDefinedName(store.workbook, store.state.activeSheet, currentRange(sel)),
		scope: undefined,
		comment: '',
		refersTo: `=${selectionReference(store.activeSheet.name, sel.ranges)}`,
	};
}

/** Edit Name fields of an existing definition. */
export function editNameDraft(def: DefinedName): NameDraft {
	return {
		name: definedNameDisplayName(def),
		scope: def.localSheet,
		comment: def.comment ?? '',
		refersTo: `=${def.formula}`,
	};
}

function valueOf(store: ViewerStore, def: DefinedName): string {
	const { row, col } = store.selectionOf().active;
	const sheet = def.localSheet ?? store.state.activeSheet;
	try {
		return definedNameDisplayValue(
			store.workbook,
			def,
			(formula, s, r, c) => store.calc.evaluateFormula(formula, s, r, c),
			{ sheet, row, col },
		);
	} catch {
		return '';
	}
}

/** Name Manager rows: visible names sorted by name then scope, built-ins without `_xlnm.`. */
export function nameManagerRows(store: ViewerStore): NameRow[] {
	const wb = store.workbook;
	return wb.definedNames
		.filter((def) => !def.hidden)
		.sort(compareDefinedNames)
		.map((def) => ({
			key: definedNameKey(def),
			def,
			name: definedNameDisplayName(def),
			value: valueOf(store, def),
			refersTo: `=${def.formula}`,
			scope:
				def.localSheet === undefined
					? store.t('dialog.names.workbook')
					: (wb.sheets[def.localSheet]?.name ?? ''),
			comment: def.comment ?? '',
		}));
}

/** Run a name edit as one undoable step; core errors (invalid / duplicate name) become messages. */
function run(
	store: ViewerStore,
	label: string,
	edit: (tx: EditTransaction) => DefinedName | undefined,
): NameCommandResult {
	let def: DefinedName | undefined;
	try {
		store.transact(label, (tx) => {
			def = edit(tx);
		});
	} catch (e) {
		return { error: (e as Error).message };
	}
	return { def };
}

/** Formulas > Define Name: add a name as one undoable step ('Define Name'). */
export function defineNameCmd(store: ViewerStore, def: DefinedName): NameCommandResult {
	return run(store, store.t('cmd.defineName'), (tx) => addDefinedName(tx, store.workbook, def));
}

function sameDefinition(a: DefinedName, b: DefinedName): boolean {
	return (
		definedNameDisplayName(a) === definedNameDisplayName(b) &&
		a.formula === b.formula &&
		(a.comment?.trim() || '') === (b.comment?.trim() || '') &&
		a.localSheet === b.localSheet
	);
}

/**
 * Name Manager > Edit, or a committed Refers to: replace a definition as
 * one undoable step ('Edit Name'). Unchanged definitions are left alone.
 */
export function editNameCmd(
	store: ViewerStore,
	target: DefinedName,
	next: DefinedName,
): NameCommandResult {
	if (sameDefinition(target, next)) {
		return { def: target };
	}
	return run(store, store.t('undo.editName'), (tx) =>
		updateDefinedName(tx, store.workbook, target, next),
	);
}

/** Name Manager > Delete: remove a definition as one undoable step. */
export function deleteNameCmd(store: ViewerStore, target: DefinedName): boolean {
	let removed = false;
	store.transact(store.t('undo.deleteName'), (tx) => {
		removed = deleteDefinedName(tx, store.workbook, target);
	});
	return removed;
}

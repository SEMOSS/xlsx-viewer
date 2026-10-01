import {
	getActiveFunctionCall,
	getCompletionPrefix,
	getFunctionInfo,
	searchFunctions,
	toggleReferenceAbsolute,
} from 'xlsx-viewer-core';
import { useMemo, useState, type KeyboardEvent } from 'react';

import { useT } from '../../../i18n';
import { buildArgHint, type ArgHint } from '../../../engine/formula/arg-hint';
import { updateEditText } from '../../commands/edit-commands';
import type { EditState } from '../../state/types';
import type { ViewerStore } from '../../state/viewer-store';

export interface Completion {
	name: string;
	kind: 'function' | 'name' | 'table';
	description?: string;
}

export interface AssistModel {
	completions: Completion[];
	highlighted: number;
	/** Where the completed identifier starts in the editor text. */
	start: number;
	/** Argument tooltip for the function call around the caret. */
	hint: ArgHint | null;
	/** Handle a key while editing; returns true when the assist consumed it. */
	onKeyDown: (e: KeyboardEvent<HTMLElement>) => boolean;
	accept: (index: number) => void;
}

const NONE = { completions: [] as Completion[], start: 0, prefix: '' };

/**
 * Formula autocomplete (functions, defined names, tables), the argument
 * tooltip and F4. Tab accepts a completion, Up/Down move through the list,
 * Escape closes it (a second Escape cancels the edit as usual).
 */
export function useFormulaAssist(
	store: ViewerStore,
	edit: EditState | null,
	active: boolean,
): AssistModel {
	const t = useT();
	const [hl, setHl] = useState({ prefix: '', index: 0 });
	const [dismissedAt, setDismissedAt] = useState<string | null>(null);
	const isFormula = active && !!edit && edit.text.startsWith('=');
	const key = edit ? `${edit.text}|${edit.caret}` : '';

	const { completions, start, prefix } = useMemo(() => {
		if (!edit || !isFormula || dismissedAt === key) {
			return NONE;
		}
		const found = getCompletionPrefix(edit.text, edit.caret);
		if (!found || found.prefix.length === 0) {
			return NONE;
		}
		const upper = found.prefix.toUpperCase();
		const list: Completion[] = searchFunctions(upper, 14).map((f) => ({
			name: f.name,
			kind: 'function',
			description: f.description,
		}));
		for (const n of store.workbook.definedNames) {
			if (
				!n.hidden &&
				!n.builtin &&
				n.name.toUpperCase().startsWith(upper) &&
				(n.localSheet === undefined || n.localSheet === edit.sheet)
			) {
				list.push({ name: n.name, kind: 'name', description: n.comment || `=${n.formula}` });
			}
		}
		for (const sheet of store.workbook.sheets) {
			for (const table of sheet.tables) {
				if (table.displayName.toUpperCase().startsWith(upper)) {
					list.push({
						name: table.displayName,
						kind: 'table',
						description: t('formulaAssist.tableOn', { sheet: sheet.name }),
					});
				}
			}
		}
		return { completions: list.slice(0, 20), start: found.start, prefix: upper };
	}, [edit, isFormula, dismissedAt, key, store, t]);

	const hint = useMemo(() => {
		if (!edit || !isFormula) {
			return null;
		}
		const call = getActiveFunctionCall(edit.text, edit.caret);
		const info = call ? getFunctionInfo(call.name) : undefined;
		return call && info ? buildArgHint(info, call.argIndex) : null;
	}, [edit, isFormula]);

	const highlighted =
		hl.prefix === prefix ? Math.min(hl.index, Math.max(0, completions.length - 1)) : 0;

	const accept = (index: number) => {
		const item = completions[index];
		if (!edit || !item) {
			return;
		}
		const insert = item.kind === 'function' ? `${item.name}(` : item.name;
		const text = edit.text.slice(0, start) + insert + edit.text.slice(edit.caret);
		updateEditText(store, text, start + insert.length);
	};

	const onKeyDown = (e: KeyboardEvent<HTMLElement>): boolean => {
		if (e.key === 'F4' && isFormula && edit && !e.altKey && !e.ctrlKey && !e.metaKey) {
			e.preventDefault();
			const next = toggleReferenceAbsolute(edit.text, edit.caret);
			if (next.formula !== edit.text || next.caret !== edit.caret) {
				store.setUi({
					edit: { ...edit, text: next.formula, caret: next.caret, point: undefined, mode: 'edit' },
				});
			}
			return true;
		}
		if (completions.length === 0) {
			return false;
		}
		switch (e.key) {
			case 'ArrowDown':
			case 'ArrowUp': {
				e.preventDefault();
				const delta = e.key === 'ArrowDown' ? 1 : -1;
				setHl({ prefix, index: (highlighted + delta + completions.length) % completions.length });
				return true;
			}
			case 'Tab':
				e.preventDefault();
				accept(highlighted);
				return true;
			case 'Escape':
				e.preventDefault();
				setDismissedAt(key);
				return true;
			default:
				return false;
		}
	};

	return { completions, highlighted, start, hint, onKeyDown, accept };
}

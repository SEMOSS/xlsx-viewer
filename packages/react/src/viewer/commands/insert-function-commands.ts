import type { FunctionInfo } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';
import { beginEdit, updateEditText } from './edit-commands';

/** Excel's initial "Most Recently Used" list. */
const DEFAULT_RECENT = [
	'SUM',
	'AVERAGE',
	'IF',
	'HYPERLINK',
	'COUNT',
	'MAX',
	'SIN',
	'SUMIF',
	'PMT',
	'STDEV',
];
const MAX_RECENT = 10;

let recent: string[] = [...DEFAULT_RECENT];

/** Functions used recently from Insert Function (most recent first). */
export function recentFunctions(): readonly string[] {
	return recent;
}

export function rememberFunction(name: string): void {
	const upper = name.toUpperCase();
	recent = [upper, ...recent.filter((n) => n !== upper)].slice(0, MAX_RECENT);
}

/** Reset the recently-used list (tests). */
export function resetRecentFunctions(): void {
	recent = [...DEFAULT_RECENT];
}

/**
 * "Search for a function": names matching the query rank first (exact,
 * prefix, contains), then functions whose description contains every word
 * of the query.
 */
export function searchFunctionCatalog(query: string, all: readonly FunctionInfo[]): FunctionInfo[] {
	const q = query.trim().toUpperCase();
	if (!q) {
		return [];
	}
	const words = q
		.toLowerCase()
		.split(/\s+/)
		.filter((w) => w.length > 1 || q.length === 1);
	const scored: Array<{ f: FunctionInfo; score: number }> = [];
	for (const f of all) {
		let score = 0;
		if (f.name === q) {
			score = 4;
		} else if (f.name.startsWith(q)) {
			score = 3;
		} else if (f.name.includes(q.replace(/\s+/g, ''))) {
			score = 2;
		} else {
			const text = `${f.name} ${f.description}`.toLowerCase();
			if (words.length && words.every((w) => text.includes(w))) {
				score = 1;
			}
		}
		if (score > 0) {
			scored.push({ f, score });
		}
	}
	return scored.sort((a, b) => b.score - a.score || (a.f.name < b.f.name ? -1 : 1)).map((s) => s.f);
}

/**
 * Start a function call in the active cell (Insert Function > OK): while a
 * formula is being edited, `NAME(` goes in at the caret; otherwise editing
 * starts with `=NAME(`.
 */
export function insertFunctionCall(store: ViewerStore, name: string): void {
	if (store.readOnly) {
		return;
	}
	const call = `${name.toUpperCase()}(`;
	const edit = store.state.edit;
	if (edit && edit.text.startsWith('=')) {
		const text = edit.text.slice(0, edit.caret) + call + edit.text.slice(edit.caret);
		updateEditText(store, text, edit.caret + call.length);
		return;
	}
	if (edit) {
		updateEditText(store, `=${call}`, call.length + 1);
		return;
	}
	beginEdit(store, { initialText: `=${call}`, mode: 'enter' });
}

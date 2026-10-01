import type { ManagedRule, RangeAddress } from 'xlsx-viewer-core';
import {
	formatRange,
	parseReferenceText,
	rangesIntersect,
	refAreaToRange,
	replaceConditionalRules,
	sameSheetName,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

/** A rule in the Rules Manager's working copy. */
export interface DraftRule extends ManagedRule {
	id: string;
}

/** "Applies to" text: `=$A$1:$A$10,$C$1:$C$5`. */
export function formatAppliesTo(ranges: readonly RangeAddress[]): string {
	return `=${ranges.map((r) => formatRange(r, { absolute: true })).join(',')}`;
}

/** Parse "Applies to" text typed in the manager; null when invalid or on another sheet. */
export function parseAppliesTo(text: string, sheetName: string): RangeAddress[] | null {
	const body = text.trim().replace(/^=/, '');
	if (!body) {
		return null;
	}
	const out: RangeAddress[] = [];
	for (const part of body.split(',')) {
		const ref = parseReferenceText(part.trim());
		if (
			!ref ||
			(ref.sheet &&
				(ref.sheet.workbook !== undefined || !sameSheetName(ref.sheet.sheet, sheetName)))
		) {
			return null;
		}
		out.push(refAreaToRange(ref.area));
	}
	return out;
}

/** Rules shown for "Current Selection" (rules touching it) or "This Worksheet" (all). */
export function visibleRules(
	rules: readonly DraftRule[],
	scope: 'selection' | 'sheet',
	selection: readonly RangeAddress[],
): DraftRule[] {
	if (scope === 'sheet') {
		return [...rules];
	}
	return rules.filter((d) => d.ranges.some((r) => selection.some((s) => rangesIntersect(r, s))));
}

/** Move a rule up (-1) or down (+1) past the next visible rule, changing its priority. */
export function moveRule(
	rules: readonly DraftRule[],
	id: string,
	delta: -1 | 1,
	visible: readonly DraftRule[],
): DraftRule[] {
	const v = visible.findIndex((d) => d.id === id);
	const neighbour = visible[v + delta];
	if (v < 0 || !neighbour) {
		return [...rules];
	}
	const from = rules.findIndex((d) => d.id === id);
	const to = rules.findIndex((d) => d.id === neighbour.id);
	const next = [...rules];
	const [moved] = next.splice(from, 1);
	next.splice(to, 0, moved);
	return next;
}

/** Duplicate a rule just below itself (lower priority), like Excel's Duplicate Rule. */
export function duplicateRule(rules: readonly DraftRule[], id: string, newId: string): DraftRule[] {
	const i = rules.findIndex((d) => d.id === id);
	if (i < 0) {
		return [...rules];
	}
	const copy: DraftRule = {
		...rules[i],
		id: newId,
		rule: { ...rules[i].rule, extId: undefined },
		ranges: rules[i].ranges.map((r) => ({ ...r })),
	};
	return [...rules.slice(0, i + 1), copy, ...rules.slice(i + 1)];
}

/** Write the manager's rules back to the active sheet as one undoable step. */
export function applyManagedRules(store: ViewerStore, rules: readonly DraftRule[]): boolean {
	const sheetIndex = store.state.activeSheet;
	return store.transact(store.t('undo.manageRules'), (tx) =>
		replaceConditionalRules(
			tx,
			store.workbook,
			sheetIndex,
			rules.map(({ id: _id, ...managed }) => managed),
		),
	);
}

import type { CfRule, RangeAddress } from 'xlsx-viewer-core';
import { addConditionalFormat, clearConditionalFormats, rangesIntersect } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

/**
 * Add a conditional formatting rule to the selected ranges as one undoable
 * step (Home > Conditional Formatting). Inline formats are interned into
 * the stylesheet's dxfs. Returns false when read-only.
 */
export function addConditionalRule(
	store: ViewerStore,
	rule: CfRule,
	label: string = store.t('cmd.conditionalFormatting'),
): boolean {
	const sheetIndex = store.state.activeSheet;
	const ranges = store.selectionOf().ranges;
	return store.transact(label, (tx) => {
		tx.touchSheet(sheetIndex);
		addConditionalFormat(store.workbook.sheets[sheetIndex], ranges, rule, store.workbook.styles);
	});
}

function touches(
	formats: ReadonlyArray<{ ranges: RangeAddress[] }>,
	ranges: readonly RangeAddress[],
): boolean {
	return formats.some((f) => f.ranges.some((r) => ranges.some((s) => rangesIntersect(r, s))));
}

/**
 * Clear Rules from Selected Cells / from Entire Sheet. Does nothing (no undo
 * step) when there are no rules to clear.
 */
export function clearConditionalRules(store: ViewerStore, scope: 'selection' | 'sheet'): boolean {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const ranges = store.selectionOf().ranges;
	const formats = sheet.conditionalFormats ?? [];
	if (formats.length === 0 || (scope === 'selection' && !touches(formats, ranges))) {
		return false;
	}
	return store.transact(store.t('undo.clearRules'), (tx) => {
		tx.touchSheet(sheetIndex);
		clearConditionalFormats(sheet, scope === 'selection' ? ranges : undefined);
	});
}

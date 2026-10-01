import type { RangeAddress, TextColumnFormat, TextSplitOptions } from 'xlsx-viewer-core';
import {
	cellRange,
	formatRange,
	guessFixedWidthBreaks,
	parseReferenceText,
	refAreaToRange,
	sameSheetName,
	TEXT_TO_COLUMNS_MESSAGES,
	textToColumnsLines,
	trimRangeToUsed,
} from 'xlsx-viewer-core';

import { currentRange } from '../../../../engine/selection/selection';
import type { Translate } from '../../../../i18n';
import type { ViewerStore } from '../../../state/viewer-store';

/** State and helpers of the Convert Text to Columns Wizard (no React). */

export type SplitMode = 'delimited' | 'fixed';

export interface DelimiterState {
	tab: boolean;
	semicolon: boolean;
	comma: boolean;
	space: boolean;
	other: boolean;
	otherChar: string;
	consecutive: boolean;
	qualifier: '"' | "'" | 'none';
}

/** Excel's first-use settings: Tab, with `"` as the text qualifier. */
export const DEFAULT_DELIMITERS: DelimiterState = {
	tab: true,
	semicolon: false,
	comma: false,
	space: false,
	other: false,
	otherChar: '',
	consecutive: false,
	qualifier: '"',
};

/** Lines shown in the previews (the conversion itself uses every line). */
export const PREVIEW_LINES = 50;

export interface TextToColumnsSetup {
	/** The one-column source, limited to the used rows; null when the wizard cannot run. */
	source: RangeAddress | null;
	/** Text of each source cell. */
	lines: string[];
	/** Excel's message when the selection cannot be converted. */
	error: string | null;
}

export function textToColumnsSetup(store: ViewerStore): TextToColumnsSetup {
	const sel = store.selectionOf();
	if (sel.ranges.length > 1) {
		return { source: null, lines: [], error: store.t('dialog.textToColumns.multipleSelections') };
	}
	const range = currentRange(sel);
	if (range.left !== range.right) {
		return { source: null, lines: [], error: TEXT_TO_COLUMNS_MESSAGES.multipleColumns };
	}
	const source = trimRangeToUsed(store.activeSheet, range);
	const lines = textToColumnsLines(store.workbook, store.state.activeSheet, source);
	if (lines.every((l) => l === '')) {
		return { source: null, lines: [], error: store.t('dialog.textToColumns.noData') };
	}
	return { source, lines, error: null };
}

/** The wizard's first guess: fixed width for space-aligned columns, else delimited. */
export function guessSplitMode(lines: readonly string[]): SplitMode {
	const used = lines.filter((l) => l !== '');
	if (used.some((l) => /[\t,;]/.test(l))) {
		return 'delimited';
	}
	return used.length > 1 && guessFixedWidthBreaks(used).length > 0 ? 'fixed' : 'delimited';
}

export function splitOptionsOf(
	mode: SplitMode,
	d: DelimiterState,
	breaks: readonly number[],
): TextSplitOptions {
	if (mode === 'fixed') {
		return { mode, breaks: [...breaks] };
	}
	return {
		mode,
		tab: d.tab,
		semicolon: d.semicolon,
		comma: d.comma,
		space: d.space,
		other: d.other && d.otherChar ? d.otherChar : undefined,
		treatConsecutiveAsOne: d.consecutive,
		textQualifier: d.qualifier === 'none' ? null : d.qualifier,
	};
}

/** `5, 12` for the accessible break list. */
export function formatBreaks(breaks: readonly number[]): string {
	return breaks.join(', ');
}

/** Break positions typed as a list (`5, 12` or `5 12`); invalid entries are ignored. */
export function parseBreaksText(text: string): number[] {
	const out = new Set<number>();
	for (const part of text.split(/[\s,;]+/)) {
		const n = Number(part);
		if (part !== '' && Number.isInteger(n) && n > 0) {
			out.add(n);
		}
	}
	return [...out].sort((a, b) => a - b);
}

/** Add a break, or ignore the click when a break is already there. */
export function addBreak(breaks: readonly number[], at: number, maxWidth: number): number[] {
	if (at <= 0 || at >= maxWidth || breaks.includes(at)) {
		return [...breaks];
	}
	return [...breaks, at].sort((a, b) => a - b);
}

/** The cell where the result starts, written like Excel (`$A$1`). */
export function formatDestination(row: number, col: number): string {
	return formatRange(cellRange(row, col), { absolute: true });
}

/** Parse the Destination box: a cell or range (its top-left cell) on the active sheet. */
export function parseDestination(
	text: string,
	sheetName: string,
): { row: number; col: number } | null {
	const parsed = parseReferenceText(text.trim().replace(/^=/, ''));
	if (!parsed) {
		return null;
	}
	const q = parsed.sheet;
	if (
		q &&
		(q.workbook !== undefined || q.sheetEnd !== undefined || !sameSheetName(q.sheet, sheetName))
	) {
		return null;
	}
	const r = refAreaToRange(parsed.area);
	return { row: r.top, col: r.left };
}

/** Number of columns the conversion produces (at least one). */
export function columnCountOf(rows: readonly (readonly string[])[]): number {
	return rows.reduce((n, r) => Math.max(n, r.length), 1);
}

/** Column header shown in the Step 3 preview. */
export function formatLabel(format: TextColumnFormat, t: Translate): string {
	if (format === 'text') {
		return t('format.text');
	}
	return t(format === 'skip' ? 'dialog.textToColumns.skipColumn' : 'format.general');
}

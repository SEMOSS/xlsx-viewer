import type {
	BorderEdgesChange,
	ColorSpec,
	FillSpec,
	HorizontalAlignment,
	RangeAddress,
	RangeBorderState,
	StylePatch,
	UnderlineStyle,
	VerticalAlignment,
	Worksheet,
} from 'xlsx-viewer-core';
import {
	applyBorderEdges,
	applyStylePatch,
	BORDER_EDGE_NAMES,
	cellKey,
	describeRangeBorders,
	keyCol,
	keyRow,
	mergeRange,
	rangesEqual,
	rangesIntersect,
	unmergeRange,
} from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import { styleIdAt } from '../../engine/layout/style-at';
import type { ViewerStore } from '../state/viewer-store';

/** A property whose value differs between the selected cells. */
export const MIXED: unique symbol = Symbol('mixed');
export type Mixable<T> = T | typeof MIXED;

/** Everything the Format Cells dialog edits, across the selection. */
export interface FormatCellsValues {
	numFmt: Mixable<string>;
	fontName: Mixable<string>;
	fontScheme: Mixable<'major' | 'minor' | undefined>;
	fontSize: Mixable<number>;
	bold: Mixable<boolean>;
	italic: Mixable<boolean>;
	/** `'none'` when not underlined. */
	underline: Mixable<UnderlineStyle>;
	/** `undefined` = Automatic. */
	fontColor: Mixable<ColorSpec | undefined>;
	strike: Mixable<boolean>;
	vertAlign: Mixable<'superscript' | 'subscript' | 'baseline'>;
	horizontal: Mixable<HorizontalAlignment>;
	vertical: Mixable<VerticalAlignment>;
	indent: Mixable<number>;
	wrapText: Mixable<boolean>;
	shrinkToFit: Mixable<boolean>;
	/** 0-90 counter-clockwise, 91-180 clockwise, 255 stacked. */
	textRotation: Mixable<number>;
	/** 0 context, 1 left-to-right, 2 right-to-left. */
	readingOrder: Mixable<number>;
	justifyLastLine: Mixable<boolean>;
	merged: Mixable<boolean>;
	/** `undefined` = no fill. */
	fill: Mixable<FillSpec | undefined>;
	locked: Mixable<boolean>;
	hidden: Mixable<boolean>;
}

export interface FormatCellsContext {
	values: FormatCellsValues;
	borders: RangeBorderState;
	/** The selection has more than one row / column (inside borders apply). */
	multiRow: boolean;
	multiCol: boolean;
	/** Value and format of the active cell (Number tab sample). */
	sampleValue: unknown;
	activeFormat: string;
}

/** Cap on distinct formats inspected for the mixed state. */
const MAX_STYLES = 400;
/** Ranges up to this many cells are inspected cell by cell. */
const SCAN_LIMIT = 20000;

/**
 * Cell formats used in the ranges: per cell for small ranges, else the
 * stored cells plus the row / column formats.
 */
export function selectionStyleIds(sheet: Worksheet, ranges: readonly RangeAddress[]): number[] {
	const ids = new Set<number>();
	for (const r of ranges) {
		const area = (r.bottom - r.top + 1) * (r.right - r.left + 1);
		if (area <= SCAN_LIMIT) {
			for (let row = r.top; row <= r.bottom && ids.size < MAX_STYLES; row++) {
				for (let col = r.left; col <= r.right; col++) {
					ids.add(styleIdAt(sheet, row, col));
				}
			}
			continue;
		}
		for (const [key, cell] of sheet.cells) {
			const row = keyRow(key);
			const col = keyCol(key);
			if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) {
				ids.add(styleIdAt(sheet, row, col, cell));
				if (ids.size >= MAX_STYLES) {
					break;
				}
			}
		}
		// Empty positions: corners and the row / column defaults.
		for (const [row, col] of [
			[r.top, r.left],
			[r.bottom, r.right],
			[r.top, r.right],
			[r.bottom, r.left],
		]) {
			ids.add(styleIdAt(sheet, row, col));
		}
	}
	return [...ids];
}

function equal(a: unknown, b: unknown): boolean {
	return a === b || JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

function uniform<S, T>(items: readonly S[], pick: (s: S) => T): Mixable<T> {
	const first = pick(items[0]);
	for (let i = 1; i < items.length; i++) {
		if (!equal(pick(items[i]), first)) {
			return MIXED;
		}
	}
	return first;
}

/** Read the selection's formats for the dialog (the active cell decides nothing on its own). */
export function readFormatCells(store: ViewerStore): FormatCellsContext {
	const sheet = store.activeSheet;
	const sel = store.selectionOf();
	const ids = selectionStyleIds(sheet, sel.ranges);
	const styles = ids.map((id) => store.styles.resolve(id));
	const s = styles;
	const merges = sheet.merges;
	const merged: Mixable<boolean> = sel.ranges.every((r) => merges.some((m) => rangesEqual(m, r)))
		? true
		: sel.ranges.some((r) => merges.some((m) => rangesIntersect(m, r)))
			? MIXED
			: false;
	const range = currentRange(sel);
	const { row, col } = sel.active;
	const active = store.styles.resolve(styleIdAt(sheet, row, col));
	return {
		values: {
			numFmt: uniform(s, (x) => x.numFmt),
			fontName: uniform(s, (x) => x.font.name),
			fontScheme: uniform(s, (x) => (x.font.scheme === 'none' ? undefined : x.font.scheme)),
			fontSize: uniform(s, (x) => x.font.size),
			bold: uniform(s, (x) => Boolean(x.font.bold)),
			italic: uniform(s, (x) => Boolean(x.font.italic)),
			underline: uniform(s, (x) => x.font.underline ?? 'none'),
			fontColor: uniform(s, (x) => (x.font.color?.auto ? undefined : x.font.color)),
			strike: uniform(s, (x) => Boolean(x.font.strike)),
			vertAlign: uniform(s, (x) => x.font.vertAlign ?? 'baseline'),
			horizontal: uniform(s, (x) => x.alignment.horizontal ?? 'general'),
			vertical: uniform(s, (x) => x.alignment.vertical ?? 'bottom'),
			indent: uniform(s, (x) => x.alignment.indent ?? 0),
			wrapText: uniform(s, (x) => Boolean(x.alignment.wrapText)),
			shrinkToFit: uniform(s, (x) => Boolean(x.alignment.shrinkToFit)),
			textRotation: uniform(s, (x) => x.alignment.textRotation ?? 0),
			readingOrder: uniform(s, (x) => x.alignment.readingOrder ?? 0),
			justifyLastLine: uniform(s, (x) => Boolean(x.alignment.justifyLastLine)),
			merged,
			fill: uniform(s, (x) => x.fill),
			locked: uniform(s, (x) => x.protection.locked),
			hidden: uniform(s, (x) => x.protection.hidden),
		},
		borders: describeRangeBorders(store.workbook, store.state.activeSheet, range),
		multiRow: sel.ranges.some((r) => r.bottom > r.top),
		multiCol: sel.ranges.some((r) => r.right > r.left),
		sampleValue: sheet.cells.get(cellKey(row, col))?.value ?? null,
		activeFormat: active.numFmt,
	};
}

/** The dialog's result: only what the user changed. */
export interface FormatCellsChanges {
	patch: StylePatch;
	borders?: BorderEdgesChange;
	merge?: boolean;
}

function changedKeys(
	initial: FormatCellsValues,
	draft: FormatCellsValues,
): Set<keyof FormatCellsValues> {
	const out = new Set<keyof FormatCellsValues>();
	for (const key of Object.keys(draft) as Array<keyof FormatCellsValues>) {
		const next = draft[key];
		if (next !== MIXED && (initial[key] === MIXED || !equal(initial[key], next))) {
			out.add(key);
		}
	}
	return out;
}

type Plain<K extends keyof FormatCellsValues> = Exclude<FormatCellsValues[K], typeof MIXED>;

/** Turn the dialog's draft into a style patch holding the changed properties only. */
export function formatCellsChanges(
	initial: FormatCellsValues,
	draft: FormatCellsValues,
	borders?: BorderEdgesChange,
): FormatCellsChanges {
	const changed = changedKeys(initial, draft);
	const when = <K extends keyof FormatCellsValues>(key: K, fn: (v: Plain<K>) => void): void => {
		if (changed.has(key)) {
			fn(draft[key] as Plain<K>);
		}
	};
	const patch: StylePatch = {};
	const font: NonNullable<StylePatch['font']> = {};
	if (changed.has('fontName') || changed.has('fontScheme')) {
		font.name = draft.fontName === MIXED ? undefined : draft.fontName;
		font.scheme = draft.fontScheme === MIXED ? undefined : draft.fontScheme;
	}
	when('fontSize', (v) => (font.size = v));
	when('bold', (v) => (font.bold = v || undefined));
	when('italic', (v) => (font.italic = v || undefined));
	when('strike', (v) => (font.strike = v || undefined));
	when('underline', (v) => (font.underline = v === 'none' ? undefined : v));
	when('fontColor', (v) => (font.color = v));
	when('vertAlign', (v) => (font.vertAlign = v === 'baseline' ? undefined : v));
	if (Object.keys(font).length) {
		patch.font = font;
	}
	const alignment: NonNullable<StylePatch['alignment']> = {};
	when('horizontal', (v) => (alignment.horizontal = v === 'general' ? undefined : v));
	when('vertical', (v) => (alignment.vertical = v === 'bottom' ? undefined : v));
	when('indent', (v) => (alignment.indent = v || undefined));
	when('wrapText', (v) => (alignment.wrapText = v || undefined));
	when('shrinkToFit', (v) => (alignment.shrinkToFit = v || undefined));
	when('textRotation', (v) => (alignment.textRotation = v || undefined));
	when('readingOrder', (v) => (alignment.readingOrder = v || undefined));
	when('justifyLastLine', (v) => (alignment.justifyLastLine = v || undefined));
	if (Object.keys(alignment).length) {
		patch.alignment = alignment;
	}
	when('numFmt', (v) => (patch.numFmt = v));
	when('fill', (v) => (patch.fill = v ?? null));
	if (changed.has('locked') || changed.has('hidden')) {
		const protection: NonNullable<StylePatch['protection']> = {};
		// Locked is the default, so only "unlocked" is stored.
		when('locked', (v) => (protection.locked = v ? undefined : false));
		when('hidden', (v) => (protection.hidden = v || undefined));
		patch.protection = protection;
	}
	const out: FormatCellsChanges = { patch };
	if (borders && BORDER_EDGE_NAMES.some((n) => borders[n] !== undefined)) {
		out.borders = borders;
	}
	when('merged', (v) => (out.merge = v));
	return out;
}

/**
 * Apply Format Cells changes to every selected range as one undoable step.
 * Returns false when nothing changed.
 */
export function applyFormatCells(store: ViewerStore, changes: FormatCellsChanges): boolean {
	const hasPatch = Object.keys(changes.patch).length > 0;
	if (!hasPatch && !changes.borders && changes.merge === undefined) {
		return false;
	}
	const wb = store.workbook;
	const sheetIndex = store.state.activeSheet;
	const ranges = store.selectionOf().ranges;
	return store.transact(store.t('undo.formatCells'), (tx) => {
		for (const r of ranges) {
			if (changes.merge === true) {
				mergeRange(tx, wb, sheetIndex, r);
			} else if (changes.merge === false) {
				unmergeRange(tx, wb, sheetIndex, r);
			}
		}
		if (hasPatch) {
			applyStylePatch(tx, wb, sheetIndex, ranges, changes.patch);
		}
		if (changes.borders) {
			for (const r of ranges) {
				applyBorderEdges(tx, wb, sheetIndex, r, changes.borders);
			}
		}
	});
}

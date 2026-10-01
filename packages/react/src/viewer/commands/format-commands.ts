import type {
	AlignmentSpec,
	BorderEdge,
	BorderPreset,
	ColorSpec,
	FontSpec,
	HorizontalAlignment,
	StylePatch,
	VerticalAlignment,
} from 'xlsx-viewer-core';
import {
	adjustDecimals,
	applyBorderPreset,
	applyStylePatch,
	cellKey,
	clearRanges,
	expandToMerges,
	mergeRange,
	numberFormatCode,
	unmergeRange,
} from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import { styleIdAt } from '../../engine/layout/style-at';
import type { TranslationKey } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';

/** Resolved style of the active cell (what the ribbon reflects). */
export function activeStyle(store: ViewerStore) {
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	return store.styles.resolve(styleIdAt(sheet, row, col));
}

function patchSelection(
	store: ViewerStore,
	label: string,
	patch: StylePatch | ((current: number | undefined) => StylePatch),
): void {
	const sheetIndex = store.state.activeSheet;
	const ranges = store.selectionOf().ranges;
	store.transact(label, (tx) => applyStylePatch(tx, store.workbook, sheetIndex, ranges, patch));
}

type FontFlag = 'bold' | 'italic' | 'strike';

const FONT_FLAG_LABELS: Record<FontFlag, TranslationKey> = {
	bold: 'cmd.bold',
	italic: 'cmd.italic',
	strike: 'cmd.strikethrough',
};

export function toggleFontFlag(store: ViewerStore, flag: FontFlag): void {
	const on = !activeStyle(store).font[flag];
	patchSelection(store, store.t(FONT_FLAG_LABELS[flag]), {
		font: { [flag]: on } as Partial<FontSpec>,
	});
}

export function toggleUnderline(store: ViewerStore, style: 'single' | 'double' = 'single'): void {
	const current = activeStyle(store).font.underline;
	const next = current && current !== 'none' && current === style ? undefined : style;
	patchSelection(store, store.t('cmd.underline'), { font: { underline: next } });
}

export function setFontName(store: ViewerStore, name: string): void {
	patchSelection(store, store.t('cmd.fontName'), { font: { name, scheme: undefined } });
}

export function setFontSize(store: ViewerStore, size: number): void {
	if (!(size >= 1 && size <= 409)) {
		return;
	}
	patchSelection(store, store.t('cmd.fontSize'), { font: { size } });
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** Increase / Decrease Font Size buttons (step through Excel's size list). */
export function stepFontSize(store: ViewerStore, direction: 1 | -1): void {
	const sheet = store.activeSheet;
	const label = store.t(direction > 0 ? 'cmd.increaseFontSize' : 'cmd.decreaseFontSize');
	patchSelection(store, label, (current) => {
		const size =
			store.workbook.styles.fonts[store.workbook.styles.cellXfs[current ?? 0]?.fontId ?? 0]?.size ??
			11;
		const next =
			direction > 0
				? (FONT_SIZES.find((s) => s > size) ?? size + 10)
				: ([...FONT_SIZES].reverse().find((s) => s < size) ?? Math.max(1, size - 1));
		void sheet;
		return { font: { size: next } };
	});
}

export function setFontColor(store: ViewerStore, color: ColorSpec | undefined): void {
	patchSelection(store, store.t('cmd.fontColor'), { font: { color } });
}

export function setFillColor(store: ViewerStore, color: ColorSpec | null): void {
	patchSelection(store, store.t('cmd.fillColor'), {
		fill: color
			? { type: 'pattern', patternType: 'solid', fgColor: color, bgColor: { indexed: 64 } }
			: null,
	});
}

export function setAlignment(
	store: ViewerStore,
	alignment: Partial<AlignmentSpec>,
	label = store.t('group.alignment'),
): void {
	patchSelection(store, label, { alignment });
}

export function toggleHorizontal(store: ViewerStore, h: HorizontalAlignment): void {
	const current = activeStyle(store).alignment.horizontal;
	setAlignment(store, { horizontal: current === h ? undefined : h });
}

export function toggleVertical(store: ViewerStore, v: VerticalAlignment): void {
	setAlignment(store, { vertical: v === 'bottom' ? undefined : v });
}

export function toggleWrapText(store: ViewerStore): void {
	setAlignment(
		store,
		{ wrapText: !activeStyle(store).alignment.wrapText },
		store.t('cmd.wrapText'),
	);
}

export function changeIndent(store: ViewerStore, delta: 1 | -1): void {
	const label = store.t(delta > 0 ? 'cmd.increaseIndent' : 'cmd.decreaseIndent');
	patchSelection(store, label, (current) => {
		const al = store.workbook.styles.cellXfs[current ?? 0]?.alignment ?? {};
		const indent = Math.max(0, Math.min(250, (al.indent ?? 0) + delta));
		const horizontal =
			indent > 0 && (!al.horizontal || al.horizontal === 'general' || al.horizontal === 'center')
				? 'left'
				: al.horizontal;
		return { alignment: { indent: indent || undefined, horizontal } };
	});
}

export function setTextRotation(store: ViewerStore, rotation: number | undefined): void {
	setAlignment(store, { textRotation: rotation }, store.t('cmd.orientation'));
}

export function setNumberFormat(store: ViewerStore, code: string): void {
	patchSelection(store, store.t('cmd.numberFormat'), { numFmt: code });
}

/** Increase / Decrease Decimal. */
export function changeDecimals(store: ViewerStore, delta: 1 | -1): void {
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	const value = sheet.cells.get(cellKey(row, col))?.value;
	const label = store.t(delta > 0 ? 'cmd.increaseDecimal' : 'cmd.decreaseDecimal');
	patchSelection(store, label, (current) => {
		const xf = store.workbook.styles.cellXfs[current ?? 0];
		const code = numberFormatCode(xf?.numFmtId ?? 0, store.workbook.styles.numFmts);
		const next = (adjustDecimals as (c: string, d: number, v?: unknown) => string)(
			code,
			delta,
			typeof value === 'number' ? value : undefined,
		);
		return { numFmt: next };
	});
}

export function applyBorders(store: ViewerStore, preset: BorderPreset, edge?: BorderEdge): void {
	const sheetIndex = store.state.activeSheet;
	const ranges = store.selectionOf().ranges;
	store.transact(store.t('cmd.borders'), (tx) => {
		for (const r of ranges) {
			applyBorderPreset(tx, store.workbook, sheetIndex, r, preset, edge);
		}
	});
}

export type MergeMode = 'center' | 'across' | 'merge' | 'unmerge';

export function mergeCells(store: ViewerStore, mode: MergeMode): void {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const range = expandToMerges(sheet.merges, currentRange(store.selectionOf()));
	const sel = store.selectionOf();
	const label = store.t(mode === 'unmerge' ? 'cmd.unmergeCells' : 'cmd.mergeCells');
	store.transact(
		label,
		(tx) => {
			if (mode === 'unmerge') {
				unmergeRange(tx, store.workbook, sheetIndex, range);
				return;
			}
			const alreadyMerged = sheet.merges.some(
				(m) =>
					m.top === range.top &&
					m.left === range.left &&
					m.bottom === range.bottom &&
					m.right === range.right,
			);
			if (alreadyMerged && mode === 'center') {
				unmergeRange(tx, store.workbook, sheetIndex, range);
				return;
			}
			mergeRange(tx, store.workbook, sheetIndex, range, {
				center: mode === 'center',
				across: mode === 'across',
			});
		},
		{
			selectionAfter: {
				active: { row: range.top, col: range.left },
				anchor: sel.anchor,
				ranges: [range],
			},
		},
	);
}

/** The active cell is the top-left of a merged area. */
export function isActiveMerged(store: ViewerStore): boolean {
	const { row, col } = store.selectionOf().active;
	return store.activeSheet.merges.some((m) => m.top === row && m.left === col);
}

export function clearFormats(store: ViewerStore): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.clearFormats'), (tx) =>
		clearRanges(tx, store.workbook, sheetIndex, store.selectionOf().ranges, 'formats'),
	);
}

type ClearKind = 'all' | 'contents' | 'formats' | 'comments' | 'hyperlinks';

const CLEAR_LABELS: Record<ClearKind, TranslationKey> = {
	all: 'cmd.clearAll',
	contents: 'cmd.clearContents',
	formats: 'cmd.clearFormats',
	comments: 'undo.clearComments',
	hyperlinks: 'cmd.clearHyperlinks',
};

export function clearAll(store: ViewerStore, kind: ClearKind): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t(CLEAR_LABELS[kind]), (tx) =>
		clearRanges(tx, store.workbook, sheetIndex, store.selectionOf().ranges, kind),
	);
}

/** Format Painter: copy the source range's formats onto `target`. */
export function paintFormats(
	store: ViewerStore,
	source: { sheet: number; range: { top: number; left: number; bottom: number; right: number } },
	target: { top: number; left: number; bottom: number; right: number },
): void {
	const wb = store.workbook;
	const src = wb.sheets[source.sheet];
	const sheetIndex = store.state.activeSheet;
	const dest = wb.sheets[sheetIndex];
	const h = source.range.bottom - source.range.top + 1;
	const w = source.range.right - source.range.left + 1;
	const height = Math.max(h, target.bottom - target.top + 1);
	const width = Math.max(w, target.right - target.left + 1);
	store.transact(store.t('cmd.formatPainter'), (tx) => {
		const area = {
			top: target.top,
			left: target.left,
			bottom: target.top + height - 1,
			right: target.left + width - 1,
		};
		tx.touchRange(sheetIndex, area, false);
		for (let r = 0; r < height; r++) {
			for (let c = 0; c < width; c++) {
				const style = styleIdAt(src, source.range.top + (r % h), source.range.left + (c % w));
				const key = cellKey(target.top + r, target.left + c);
				const cell = dest.cells.get(key);
				if (cell) {
					dest.cells.set(key, { ...cell, style: style || undefined });
				} else if (style) {
					dest.cells.set(key, { value: null, style });
				}
			}
		}
		dest.dirty = true;
	});
}

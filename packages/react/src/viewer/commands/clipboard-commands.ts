import type { Cell, RangeAddress, Worksheet } from 'xlsx-viewer-core';
import {
	cellKey,
	expandToMerges,
	formatValue,
	isCellError,
	resolveColorCss,
	setCellInput,
	translateFormula,
	usedRange,
} from 'xlsx-viewer-core';

import {
	parseHtmlTable,
	parseTsv,
	toHtmlTable,
	toTsv,
	type HtmlCell,
} from '../../engine/clipboard/serialize';
import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

export interface ClipboardPayload {
	sheetId: string;
	range: RangeAddress;
	cells: Array<Array<Cell | undefined>>;
	merges: RangeAddress[];
	cut: boolean;
	text: string;
}

let internal: ClipboardPayload | null = null;

export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'transpose' | 'noBorders';

function clampRange(sheet: Worksheet, r: RangeAddress): RangeAddress {
	const used = usedRange(sheet, true) ?? { top: 0, left: 0, bottom: 0, right: 0 };
	return {
		top: r.top,
		left: r.left,
		bottom: Math.min(r.bottom, Math.max(r.top, used.bottom)),
		right: Math.min(r.right, Math.max(r.left, used.right)),
	};
}

function displayText(store: ViewerStore, cell: Cell | undefined): string {
	if (!cell || cell.value === null) {
		return '';
	}
	const v = cell.value;
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (isCellError(v)) {
		return v.code;
	}
	return formatValue(v, store.styles.resolve(cell.style).numFmt, {
		date1904: store.workbook.date1904,
	}).text;
}

function htmlStyle(store: ViewerStore, cell: Cell | undefined): string | undefined {
	if (!cell?.style) {
		return undefined;
	}
	const s = store.styles.resolve(cell.style);
	const ctx = { theme: store.workbook.theme, indexedColors: store.workbook.styles.indexedColors };
	const parts: string[] = [];
	if (s.font.bold) parts.push('font-weight:700');
	if (s.font.italic) parts.push('font-style:italic');
	if (s.font.underline && s.font.underline !== 'none') parts.push('text-decoration:underline');
	const color = resolveColorCss(s.font.color, ctx);
	if (color) parts.push(`color:${color}`);
	if (s.fill?.type === 'pattern' && s.fill.patternType === 'solid') {
		const bg = resolveColorCss(s.fill.fgColor, ctx);
		if (bg) parts.push(`background:${bg}`);
	}
	if (s.alignment.horizontal && s.alignment.horizontal !== 'general')
		parts.push(`text-align:${s.alignment.horizontal}`);
	parts.push(`font-family:${s.font.name}`, `font-size:${s.font.size}pt`);
	return parts.join(';');
}

/** Copy (or cut) the current range; returns the data for the system clipboard. */
export function copySelection(store: ViewerStore, cut = false): { text: string; html: string } {
	const sheet = store.activeSheet;
	const range = clampRange(sheet, expandToMerges(sheet.merges, currentRange(store.selectionOf())));
	const cells: Array<Array<Cell | undefined>> = [];
	const textRows: string[][] = [];
	const htmlRows: HtmlCell[][] = [];
	const merges = sheet.merges
		.filter(
			(m) =>
				m.top >= range.top &&
				m.bottom <= range.bottom &&
				m.left >= range.left &&
				m.right <= range.right,
		)
		.map((m) => ({
			top: m.top - range.top,
			left: m.left - range.left,
			bottom: m.bottom - range.top,
			right: m.right - range.left,
		}));
	for (let r = range.top; r <= range.bottom; r++) {
		const row: Array<Cell | undefined> = [];
		const tr: string[] = [];
		const hr: HtmlCell[] = [];
		for (let c = range.left; c <= range.right; c++) {
			const cell = sheet.cells.get(cellKey(r, c));
			row.push(cell);
			const text = displayText(store, cell);
			tr.push(text);
			hr.push({ text, style: htmlStyle(store, cell) });
		}
		cells.push(row);
		textRows.push(tr);
		htmlRows.push(hr);
	}
	const text = toTsv(textRows);
	internal = { sheetId: sheet.id, range, cells, merges, cut, text };
	store.setUi({ clipboard: { sheetId: sheet.id, ranges: [range], cut } });
	return { text, html: toHtmlTable(htmlRows) };
}

function normalize(text: string): string {
	return text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
}

/** Paste system clipboard data (or the internal clipboard when it is the same content). */
export function pasteData(
	store: ViewerStore,
	data: { text?: string; html?: string },
	mode: PasteMode = 'all',
): void {
	if (store.readOnly) {
		return;
	}
	if (internal && (data.text === undefined || normalize(data.text) === normalize(internal.text))) {
		pasteInternal(store, internal, mode);
		return;
	}
	const rows =
		(data.html ? parseHtmlTable(data.html) : null) ??
		(data.text !== undefined ? parseTsv(data.text) : null);
	if (!rows || rows.length === 0) {
		return;
	}
	const sheetIndex = store.state.activeSheet;
	const at = currentRange(store.selectionOf());
	const height = rows.length;
	const width = Math.max(...rows.map((r) => r.length));
	const target = {
		top: at.top,
		left: at.left,
		bottom: at.top + height - 1,
		right: at.left + width - 1,
	};
	store.transact(
		store.t('cmd.paste'),
		(tx) => {
			rows.forEach((row, r) =>
				row.forEach((text, c) => {
					setCellInput(tx, store.workbook, sheetIndex, at.top + r, at.left + c, text);
				}),
			);
		},
		{
			selectionAfter: {
				active: { row: target.top, col: target.left },
				anchor: { row: target.top, col: target.left },
				ranges: [target],
			},
		},
	);
}

function pasteInternal(store: ViewerStore, clip: ClipboardPayload, mode: PasteMode): void {
	const wb = store.workbook;
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const sourceIndex = wb.sheets.findIndex((s) => s.id === clip.sheetId);
	const at = currentRange(store.selectionOf());
	const srcH = clip.cells.length;
	const srcW = clip.cells[0]?.length ?? 0;
	const transpose = mode === 'transpose';
	const outH = transpose ? srcW : srcH;
	const outW = transpose ? srcH : srcW;
	// Tile the copy over a selection that is an exact multiple of its size (Excel behaviour).
	const selH = at.bottom - at.top + 1;
	const selW = at.right - at.left + 1;
	const tilesY = !clip.cut && selH % outH === 0 ? selH / outH : 1;
	const tilesX = !clip.cut && selW % outW === 0 ? selW / outW : 1;
	const target = {
		top: at.top,
		left: at.left,
		bottom: at.top + outH * tilesY - 1,
		right: at.left + outW * tilesX - 1,
	};

	store.transact(
		store.t(clip.cut ? 'undo.cutPaste' : 'cmd.paste'),
		(tx) => {
			tx.touchRange(sheetIndex, target);
			if (clip.cut && sourceIndex >= 0) {
				tx.touchRange(sourceIndex, clip.range);
				const src = wb.sheets[sourceIndex];
				for (let r = clip.range.top; r <= clip.range.bottom; r++) {
					for (let c = clip.range.left; c <= clip.range.right; c++) {
						src.cells.delete(cellKey(r, c));
					}
				}
				src.dirty = true;
			}
			if (mode === 'all' || mode === 'formats' || transpose) {
				tx.touchSheet(sheetIndex);
				sheet.merges = sheet.merges.filter(
					(m) =>
						m.bottom < target.top ||
						m.top > target.bottom ||
						m.right < target.left ||
						m.left > target.right,
				);
			}
			for (let ty = 0; ty < tilesY; ty++) {
				for (let tx2 = 0; tx2 < tilesX; tx2++) {
					for (let r = 0; r < srcH; r++) {
						for (let c = 0; c < srcW; c++) {
							const src = clip.cells[r][c];
							const outR = at.top + ty * outH + (transpose ? c : r);
							const outC = at.left + tx2 * outW + (transpose ? r : c);
							const key = cellKey(outR, outC);
							const existing = sheet.cells.get(key);
							let next: Cell | undefined;
							const dRow = outR - (clip.range.top + r);
							const dCol = outC - (clip.range.left + c);
							switch (mode) {
								case 'values':
									next =
										src && src.value !== null
											? { value: src.value, style: existing?.style }
											: existing?.style
												? { value: null, style: existing.style }
												: undefined;
									break;
								case 'formulas':
									next = src
										? {
												value: src.value,
												formula:
													src.formula !== undefined && !clip.cut
														? translateFormula(src.formula, dRow, dCol)
														: src.formula,
												array:
													src.formula !== undefined
														? {
																ref: { top: outR, left: outC, bottom: outR, right: outC },
																dynamic: true,
															}
														: undefined,
												style: existing?.style,
											}
										: existing?.style
											? { value: null, style: existing.style }
											: undefined;
									break;
								case 'formats':
									next = src?.style
										? { ...(existing ?? { value: null }), style: src.style }
										: existing
											? { ...existing, style: undefined }
											: undefined;
									break;
								default:
									next = src
										? {
												...src,
												arrayAnchor: undefined,
												array:
													src.formula !== undefined
														? {
																ref: { top: outR, left: outC, bottom: outR, right: outC },
																dynamic: true,
															}
														: undefined,
												formula:
													src.formula !== undefined && !clip.cut
														? translateFormula(src.formula, dRow, dCol)
														: src.formula,
											}
										: undefined;
							}
							if (next && (next.value !== null || next.formula !== undefined || next.style)) {
								for (const k of Object.keys(next) as Array<keyof Cell>) {
									if (next[k] === undefined) {
										delete next[k];
									}
								}
								sheet.cells.set(key, next);
							} else {
								sheet.cells.delete(key);
							}
						}
					}
					if (mode === 'all' && !transpose) {
						for (const m of clip.merges) {
							sheet.merges.push({
								top: at.top + ty * outH + m.top,
								left: at.left + tx2 * outW + m.left,
								bottom: at.top + ty * outH + m.bottom,
								right: at.left + tx2 * outW + m.right,
							});
						}
					}
				}
			}
			sheet.dirty = true;
		},
		{
			selectionAfter: {
				active: { row: target.top, col: target.left },
				anchor: { row: target.top, col: target.left },
				ranges: [target],
			},
		},
	);
	if (clip.cut) {
		internal = null;
		store.setUi({ clipboard: null });
	}
}

/** Clipboard API path used by ribbon buttons (no paste event available). */
export async function pasteFromSystem(store: ViewerStore, mode: PasteMode = 'all'): Promise<void> {
	if (mode !== 'all' && internal) {
		pasteInternal(store, internal, mode);
		return;
	}
	try {
		const nav = navigator.clipboard;
		let html: string | undefined;
		let text: string | undefined;
		if (nav?.read) {
			const items = await nav.read();
			for (const item of items) {
				if (item.types.includes('text/html')) {
					html = await (await item.getType('text/html')).text();
				}
				if (item.types.includes('text/plain')) {
					text = await (await item.getType('text/plain')).text();
				}
			}
		} else if (nav?.readText) {
			text = await nav.readText();
		}
		pasteData(store, { text, html }, mode);
	} catch {
		if (internal) {
			pasteInternal(store, internal, mode);
		}
	}
}

export async function copyToSystem(store: ViewerStore, cut = false): Promise<void> {
	const { text, html } = copySelection(store, cut);
	try {
		if (navigator.clipboard?.write && typeof ClipboardItem !== 'undefined') {
			await navigator.clipboard.write([
				new ClipboardItem({
					'text/plain': new Blob([text], { type: 'text/plain' }),
					'text/html': new Blob([html], { type: 'text/html' }),
				}),
			]);
		} else {
			await navigator.clipboard?.writeText(text);
		}
	} catch {
		// Internal clipboard still works inside the viewer.
	}
}

export function hasInternalClipboard(): boolean {
	return internal !== null;
}

export function clearClipboardMark(store: ViewerStore): void {
	if (store.state.clipboard) {
		store.setUi({ clipboard: null });
	}
}

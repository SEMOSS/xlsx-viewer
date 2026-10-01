/**
 * Print settings that live outside `<pageSetup>`: the fit-to-page flag of
 * `<sheetPr><pageSetUpPr>` and the manual page breaks (`<rowBreaks>`,
 * `<colBreaks>`).
 */
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import type { Worksheet } from '../types/sheet';
import { cloneElement, el } from '../xml';
import type { XmlElement } from '../xml/types';
import { findChild, findChildren, flag, mergeAttrs, setOrderedChild } from './xml-utils';

/**
 * Set `<pageSetUpPr fitToPage>` on a `<sheetPr>` being built. An undefined
 * model flag keeps the original element; sheets created in memory get the
 * flag when they carry fit-to dimensions.
 */
export function applyFitToPage(
	sheetPr: XmlElement,
	sheet: Worksheet,
	isNew: boolean,
	order: readonly string[],
	p: string,
): void {
	const setup = sheet.pageSetup;
	const fit =
		setup?.fitToPage ??
		(isNew && setup && (setup.fitToWidth !== undefined || setup.fitToHeight !== undefined)
			? true
			: undefined);
	if (fit === undefined) {
		return;
	}
	const prev = findChild(sheetPr, 'pageSetUpPr', p);
	const attrs = mergeAttrs(prev?.attrs, { fitToPage: flag(fit) });
	const next = Object.keys(attrs).length ? el('pageSetUpPr', attrs) : null;
	setOrderedChild(sheetPr, 'pageSetUpPr', next, order, p);
}

function isManual(brk: XmlElement): boolean {
	const v = brk.attrs.man;
	return v === '1' || v === 'true';
}

/**
 * `<rowBreaks>` / `<colBreaks>` from the model's manual breaks. Returns
 * `undefined` (keep the original element) when the model does not set the
 * list. Original `<brk>` attributes of kept breaks and automatic breaks of
 * the source survive; new breaks span the whole sheet like Excel's.
 */
export function buildBreaks(
	kind: 'rowBreaks' | 'colBreaks',
	orig: XmlElement | undefined,
	ids: readonly number[] | undefined,
	p: string,
): XmlElement | null | undefined {
	if (ids === undefined) {
		return undefined;
	}
	const max = kind === 'rowBreaks' ? MAX_COL_INDEX : MAX_ROW_INDEX;
	const before = findChildren(orig, 'brk', p);
	const byId = new Map<number, XmlElement>();
	const kept: Array<{ id: number; el: XmlElement }> = [];
	for (const b of before) {
		const id = Number(b.attrs.id);
		if (!Number.isFinite(id)) {
			continue;
		}
		if (isManual(b)) {
			byId.set(id, b);
		} else {
			kept.push({ id, el: cloneElement(b) });
		}
	}
	const wanted = [...new Set(ids)].filter((id) => Number.isInteger(id) && id > 0);
	for (const id of wanted) {
		const prev = byId.get(id);
		kept.push({
			id,
			el: prev ? cloneElement(prev) : el('brk', { id, max, man: 1 }),
		});
	}
	if (kept.length === 0) {
		return null;
	}
	kept.sort((a, b) => a.id - b.id);
	const attrs = mergeAttrs(orig?.attrs, {
		count: kept.length,
		manualBreakCount: wanted.length,
	});
	return el(
		kind,
		attrs,
		kept.map((k) => k.el),
	);
}

import type { Cell, Hyperlink, HyperlinkSpec, RangeAddress, Worksheet } from 'xlsx-viewer-core';
import {
	buildMailtoAddress,
	cellKey,
	compareDefinedNames,
	definedNameDisplayName,
	expandToMerges,
	findSheetIndex,
	formatHyperlinkLocation,
	formatRange,
	hyperlinkDisplayText,
	keyCol,
	keyRow,
	normalizeHyperlinkAddress,
	parseHyperlinkLocation,
	parseMailtoAddress,
	parseRange,
	rangeContains,
	setHyperlink,
} from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import type { TranslationKey } from '../../i18n';
import type { ViewerStore } from '../state/viewer-store';
import { hyperlinkForCell } from './hyperlink-commands';

/**
 * Insert / Edit Hyperlink (Ctrl+K): what the dialog starts with, the
 * places of "Place in This Document", and turning the form into a link.
 */

export type LinkKind = 'file' | 'place' | 'email';

/** An entry of "Or select a place in this document". */
export interface PlaceOption {
	key: string;
	group: 'sheet' | 'name';
	label: string;
	/** The sheet name, or the defined name used as the location. */
	value: string;
}

export interface HyperlinkForm {
	kind: LinkKind;
	/** Text to display; null when the selection holds several values. */
	text: string | null;
	tooltip: string;
	address: string;
	cellRef: string;
	/** Key of the selected place. */
	place: string;
	email: string;
	subject: string;
}

export interface HyperlinkDialogSetup {
	/** The active cell already has a link (Edit Hyperlink). */
	editing: boolean;
	sheetIndex: number;
	/** The selection, expanded to merged areas. */
	range: RangeAddress;
	/** Text to display follows the address while the user has not typed one (empty cell). */
	mirror: boolean;
	places: PlaceOption[];
	form: HyperlinkForm;
}

/** More than one cell of the range holds a value. */
function holdsSeveralValues(sheet: Worksheet, range: RangeAddress): boolean {
	let count = 0;
	const counts = (cell: Cell | undefined): boolean => {
		if (cell && cell.value !== null && cell.value !== '') {
			count++;
		}
		return count > 1;
	};
	const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
	if (area <= sheet.cells.size) {
		for (let row = range.top; row <= range.bottom; row++) {
			for (let col = range.left; col <= range.right; col++) {
				if (counts(sheet.cells.get(cellKey(row, col)))) {
					return true;
				}
			}
		}
		return false;
	}
	for (const [key, cell] of sheet.cells) {
		if (rangeContains(range, keyRow(key), keyCol(key)) && counts(cell)) {
			return true;
		}
	}
	return false;
}

const sheetKey = (index: number) => `sheet:${index}`;

/** Visible sheets, then workbook-level and active-sheet names (not hidden, not built-in). */
export function hyperlinkPlaces(store: ViewerStore): PlaceOption[] {
	const wb = store.workbook;
	const active = store.state.activeSheet;
	const out: PlaceOption[] = [];
	wb.sheets.forEach((sheet, index) => {
		if (sheet.state === 'visible') {
			out.push({ key: sheetKey(index), group: 'sheet', label: sheet.name, value: sheet.name });
		}
	});
	const names = wb.definedNames
		.filter(
			(n) => !n.hidden && !n.builtin && (n.localSheet === undefined || n.localSheet === active),
		)
		.sort(compareDefinedNames);
	for (const n of names) {
		const label = definedNameDisplayName(n);
		out.push({
			key: `name:${n.name.toLowerCase()}|${n.localSheet ?? ''}`,
			group: 'name',
			label,
			value: label,
		});
	}
	return out;
}

function formFromLink(
	link: Hyperlink,
	form: HyperlinkForm,
	places: readonly PlaceOption[],
	store: ViewerStore,
): void {
	form.tooltip = link.tooltip ?? '';
	const mail = link.target ? parseMailtoAddress(link.target) : null;
	if (mail) {
		form.kind = 'email';
		form.email = mail.email;
		form.subject = mail.subject;
	} else if (link.target) {
		form.kind = 'file';
		form.address = link.location ? `${link.target}#${link.location}` : link.target;
	} else if (link.location) {
		form.kind = 'place';
		const loc = parseHyperlinkLocation(link.location);
		const name = loc.sheet
			? undefined
			: places.find((p) => p.group === 'name' && p.value.toLowerCase() === loc.ref.toLowerCase());
		if (name) {
			form.place = name.key;
		} else {
			const index = loc.sheet ? findSheetIndex(store.workbook, loc.sheet) : -1;
			if (index >= 0 && places.some((p) => p.key === sheetKey(index))) {
				form.place = sheetKey(index);
			}
			form.cellRef = loc.ref;
		}
	}
}

/** What Insert Hyperlink starts with for the current selection. */
export function hyperlinkDialogSetup(store: ViewerStore): HyperlinkDialogSetup {
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const sel = store.selectionOf();
	const range = expandToMerges(sheet.merges, currentRange(sel));
	const link = hyperlinkForCell(store, sel.active.row, sel.active.col);
	const places = hyperlinkPlaces(store);
	const several = holdsSeveralValues(sheet, range);
	const text = several
		? null
		: hyperlinkDisplayText(store.workbook, sheetIndex, range.top, range.left);
	const form: HyperlinkForm = {
		kind: 'file',
		text,
		tooltip: '',
		address: '',
		cellRef: 'A1',
		place: places.some((p) => p.key === sheetKey(sheetIndex))
			? sheetKey(sheetIndex)
			: (places[0]?.key ?? ''),
		email: '',
		subject: '',
	};
	if (link) {
		formFromLink(link, form, places, store);
	}
	return { editing: link !== undefined, sheetIndex, range, mirror: text === '', places, form };
}

/** The text a link shows when Text to display is left empty (also mirrored while typing). */
export function defaultLinkText(form: HyperlinkForm, places: readonly PlaceOption[]): string {
	switch (form.kind) {
		case 'file':
			return form.address.trim();
		case 'email':
			return form.email.trim() === '' ? '' : buildMailtoAddress(form.email, form.subject);
		case 'place': {
			const place = places.find((p) => p.key === form.place);
			if (!place || place.group === 'name') {
				return place?.label ?? '';
			}
			const range = parseRange(form.cellRef.trim());
			return formatHyperlinkLocation(place.value, range ? formatRange(range) : form.cellRef.trim());
		}
	}
}

export type BuiltHyperlink =
	| { link: HyperlinkSpec; display?: string }
	| { error: TranslationKey }
	| null;

/**
 * The link described by the form; null when nothing was entered, `error`
 * (the key of the message to show) for a bad reference.
 */
export function buildHyperlink(
	form: HyperlinkForm,
	places: readonly PlaceOption[],
): BuiltHyperlink {
	let link: HyperlinkSpec;
	if (form.kind === 'file') {
		const address = form.address.trim();
		if (address === '') {
			return null;
		}
		link = address.startsWith('#')
			? { location: address.slice(1) }
			: { target: normalizeHyperlinkAddress(address) };
	} else if (form.kind === 'email') {
		if (form.email.trim() === '') {
			return null;
		}
		link = { target: buildMailtoAddress(form.email, form.subject) };
	} else {
		const place = places.find((p) => p.key === form.place);
		if (!place) {
			return null;
		}
		if (place.group === 'name') {
			link = { location: place.value };
		} else {
			const range = parseRange(form.cellRef.trim());
			if (!range) {
				return { error: 'message.invalidReference' };
			}
			link = { location: formatHyperlinkLocation(place.value, formatRange(range)) };
		}
	}
	const tooltip = form.tooltip.trim();
	if (tooltip !== '') {
		link.tooltip = tooltip;
	}
	if (form.text === null) {
		return { link };
	}
	const text = form.text.trim() === '' ? defaultLinkText(form, places) : form.text;
	return { link, display: text };
}

/**
 * OK: link the selection in one undoable step ('Insert Hyperlink' or
 * 'Edit Hyperlink'). Returns an error message to show, or null when done.
 */
export function applyHyperlinkForm(
	store: ViewerStore,
	setup: HyperlinkDialogSetup,
	form: HyperlinkForm,
): string | null {
	const built = buildHyperlink(form, setup.places);
	if (built === null) {
		return null;
	}
	if ('error' in built) {
		return store.t(built.error);
	}
	const label = store.t(setup.editing ? 'undo.editHyperlink' : 'undo.insertHyperlink');
	store.transact(label, (tx) => {
		setHyperlink(tx, store.workbook, setup.sheetIndex, setup.range, built.link, built.display);
	});
	return null;
}

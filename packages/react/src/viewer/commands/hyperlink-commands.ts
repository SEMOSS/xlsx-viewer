import type { Hyperlink } from 'xlsx-viewer-core';
import { findMerge, hyperlinkAt, removeHyperlinks, resolveHyperlink } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';
import { goToReference } from './selection-commands';

/** The hyperlink of a cell (or of the merged area it belongs to) on the active sheet. */
export function hyperlinkForCell(
	store: ViewerStore,
	row: number,
	col: number,
): Hyperlink | undefined {
	const sheet = store.activeSheet;
	const merge = findMerge(sheet, row, col);
	return (
		hyperlinkAt(sheet, row, col) ?? (merge ? hyperlinkAt(sheet, merge.top, merge.left) : undefined)
	);
}

function openInNewTab(url: string): void {
	window.open(url, '_blank', 'noopener,noreferrer');
}

/**
 * Follow a hyperlink: jump to an internal location (`Sheet2!A1`, a defined
 * name) or open a web or mail link in a new tab. Links to local files and
 * unsafe schemes are refused with a message.
 */
export function followHyperlink(
	store: ViewerStore,
	link: Hyperlink,
	open: (url: string) => void = openInNewTab,
): boolean {
	const dest = resolveHyperlink(link);
	if (!dest) {
		return false;
	}
	switch (dest.kind) {
		case 'location':
			if (!goToReference(store, dest.location)) {
				store.setUi({ message: store.t('message.invalidReferenceAt', { ref: dest.location }) });
				return false;
			}
			return true;
		case 'url':
			open(dest.url);
			return true;
		case 'unsupported':
			store.setUi({ message: store.t('message.linkUnsupported', { target: dest.target }) });
			return false;
	}
}

/** Remove the hyperlinks of the selected cells; cells in the Hyperlink style return to Normal. */
export function removeHyperlinksCmd(store: ViewerStore): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('undo.removeHyperlinks'), (tx) => {
		removeHyperlinks(tx, store.workbook, sheetIndex, store.selectionOf().ranges);
	});
}

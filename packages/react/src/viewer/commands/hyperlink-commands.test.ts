import { createWorkbook } from 'xlsx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ViewerStore } from '../state/viewer-store';
import { followHyperlink, hyperlinkForCell, removeHyperlinksCmd } from './hyperlink-commands';

function storeWithLinks() {
	const wb = createWorkbook({ sheetNames: ['Summary', 'Data Sheet'] });
	const sheet = wb.sheets[0];
	sheet.merges.push({ top: 4, left: 0, bottom: 4, right: 2 });
	sheet.hyperlinks.push(
		{ ref: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/docs' },
		{ ref: { top: 1, left: 0, bottom: 1, right: 0 }, location: "'Data Sheet'!C7" },
		{ ref: { top: 2, left: 0, bottom: 2, right: 0 }, target: 'javascript:alert(1)' },
		{ ref: { top: 4, left: 0, bottom: 4, right: 0 }, location: 'Summary!A1' },
	);
	return new ViewerStore(wb);
}

describe('hyperlink commands', () => {
	it('finds links on merged cells through the top-left cell', () => {
		const store = storeWithLinks();
		expect(hyperlinkForCell(store, 4, 2)?.location).toBe('Summary!A1');
		expect(hyperlinkForCell(store, 3, 0)).toBeUndefined();
	});

	it('opens web links with the opener', () => {
		const store = storeWithLinks();
		const open = vi.fn();
		expect(followHyperlink(store, hyperlinkForCell(store, 0, 0)!, open)).toBe(true);
		expect(open).toHaveBeenCalledWith('https://example.com/docs');
	});

	it('jumps to internal locations', () => {
		const store = storeWithLinks();
		expect(followHyperlink(store, hyperlinkForCell(store, 1, 0)!)).toBe(true);
		expect(store.state.activeSheet).toBe(1);
		expect(store.selectionOf(1).active).toEqual({ row: 6, col: 2 });
	});

	it('refuses unsafe schemes with a message', () => {
		const store = storeWithLinks();
		const open = vi.fn();
		expect(followHyperlink(store, hyperlinkForCell(store, 2, 0)!, open)).toBe(false);
		expect(open).not.toHaveBeenCalled();
		expect(store.state.message).toContain("can't be opened");
	});

	it('removes the links of the selection with undo', () => {
		const store = storeWithLinks();
		store.setSelection({
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			ranges: [{ top: 0, left: 0, bottom: 1, right: 0 }],
		});
		removeHyperlinksCmd(store);
		expect(store.activeSheet.hyperlinks).toHaveLength(2);
		store.undo();
		expect(store.activeSheet.hyperlinks).toHaveLength(4);
	});
});

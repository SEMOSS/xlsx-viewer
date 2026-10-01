import { createWorkbook, EditTransaction, setCellInput } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStore } from '../state/viewer-store';
import { beginEdit, commitEdit, typeInCell } from './edit-commands';

function setup() {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	['Region', 'North', 'South', 'Northeast'].forEach((v, r) => setCellInput(tx, wb, 0, r, 0, v));
	tx.commit();
	const store = new ViewerStore(wb);
	store.setSelection({
		active: { row: 4, col: 0 },
		anchor: { row: 4, col: 0 },
		ranges: [{ top: 4, left: 0, bottom: 4, right: 0 }],
	});
	return store;
}

describe('AutoComplete while typing', () => {
	it('suggests the single matching entry and commits it in its own case', () => {
		const store = setup();
		beginEdit(store, { initialText: 's' });
		expect(store.state.edit?.suggestion).toBe('South');
		expect(commitEdit(store)).toBe(true);
		expect(store.activeSheet.cells.get(4 * 16384)?.value).toBe('South');
	});

	it('waits for a unique match and drops the suggestion on delete', () => {
		const store = setup();
		beginEdit(store, { initialText: 'N' });
		expect(store.state.edit?.suggestion).toBeUndefined();
		typeInCell(store, 'Northe', 6);
		expect(store.state.edit?.suggestion).toBe('Northeast');
		// Backspace removes the selected completion: the text does not grow, no suggestion.
		typeInCell(store, 'Northe', 6);
		expect(store.state.edit?.suggestion).toBeUndefined();
		commitEdit(store);
		expect(store.activeSheet.cells.get(4 * 16384)?.value).toBe('Northe');
	});
});

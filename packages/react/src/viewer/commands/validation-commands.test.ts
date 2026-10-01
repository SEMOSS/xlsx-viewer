import { createWorkbook } from 'xlsx-viewer-core';
import type { DataValidation } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { ViewerStore } from '../state/viewer-store';
import { beginEdit, commitEdit, updateEditText } from './edit-commands';

function setup(rule: Partial<DataValidation>) {
	const wb = createWorkbook();
	wb.sheets[0].dataValidations.push({
		ranges: [{ top: 0, left: 0, bottom: 9, right: 0 }],
		type: 'whole',
		operator: 'between',
		formula1: '1',
		formula2: '10',
		showErrorMessage: true,
		...rule,
	} as DataValidation);
	return new ViewerStore(wb);
}

const type = (store: ViewerStore, text: string) => {
	beginEdit(store, { initialText: text.slice(0, 1) });
	updateEditText(store, text, text.length);
};

describe('data validation on entry', () => {
	it('accepts valid entries', () => {
		const store = setup({});
		type(store, '7');
		expect(commitEdit(store)).toBe(true);
		expect(store.activeSheet.cells.get(0)?.value).toBe(7);
	});

	it('stops invalid entries with an alert and keeps editing', () => {
		const store = setup({ errorStyle: 'stop', error: 'Between 1 and 10 please' });
		type(store, '42');
		expect(commitEdit(store)).toBe(false);
		expect(store.state.dialog).toMatchObject({
			kind: 'validationAlert',
			payload: { style: 'stop', message: 'Between 1 and 10 please' },
		});
		expect(store.state.edit?.text).toBe('42');
		expect(store.activeSheet.cells.get(0)).toBeUndefined();
	});

	it('lets a warning be accepted', () => {
		const store = setup({ errorStyle: 'warning' });
		type(store, '42');
		commitEdit(store);
		expect(store.state.dialog?.kind).toBe('validationAlert');
		expect(commitEdit(store, { skipValidation: true })).toBe(true);
		expect(store.activeSheet.cells.get(0)?.value).toBe(42);
	});

	it('accepts anything when the rule has no error alert', () => {
		const store = setup({ showErrorMessage: false });
		type(store, '42');
		expect(commitEdit(store)).toBe(true);
	});
});

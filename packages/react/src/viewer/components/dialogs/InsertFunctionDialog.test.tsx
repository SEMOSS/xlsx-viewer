import { fireEvent, screen, within } from '@testing-library/react';
import { beforeEach, describe, expect, it } from 'vitest';
import { listFunctions } from 'xlsx-viewer-core';

import {
	insertFunctionCall,
	recentFunctions,
	resetRecentFunctions,
	searchFunctionCatalog,
} from '../../commands/insert-function-commands';
import { beginEdit } from '../../commands/edit-commands';
import { renderDialog } from './__tests__/render-dialog';
import { InsertFunctionDialog } from './InsertFunctionDialog';

beforeEach(() => resetRecentFunctions());

describe('searchFunctionCatalog', () => {
	it('ranks name matches before description matches', () => {
		const all = listFunctions();
		const hits = searchFunctionCatalog('sum', all).map((f) => f.name);
		expect(hits[0]).toBe('SUM');
		expect(hits.indexOf('SUMIF')).toBeGreaterThan(0);
		expect(searchFunctionCatalog('', all)).toEqual([]);
		const byText = searchFunctionCatalog('standard deviation', all).map((f) => f.name);
		expect(byText).toContain('STDEV.S');
	});
});

describe('insertFunctionCall', () => {
	it('starts editing the active cell with =NAME(', () => {
		const { store } = renderDialog(InsertFunctionDialog);
		insertFunctionCall(store, 'vlookup');
		expect(store.state.edit).toMatchObject({ text: '=VLOOKUP(', caret: 9, row: 0, col: 0 });
	});

	it('inserts NAME( at the caret of a formula being edited', () => {
		const { store } = renderDialog(InsertFunctionDialog);
		beginEdit(store, { initialText: '=1+', mode: 'enter' });
		insertFunctionCall(store, 'SUM');
		expect(store.state.edit).toMatchObject({ text: '=1+SUM(', caret: 7 });
	});
});

describe('InsertFunctionDialog', () => {
	it('opens on Most Recently Used with syntax and description of the selection', () => {
		renderDialog(InsertFunctionDialog);
		const list = screen.getByRole('listbox', { name: 'Select a function' });
		expect(
			within(list)
				.getAllByRole('option')
				.map((o) => o.textContent),
		).toEqual([...recentFunctions()]);
		expect(screen.getByTestId('function-syntax').textContent).toBe('SUM(number1, [number2], ...)');
	});

	it('filters by category and inserts the chosen function on OK', () => {
		const { store, onClose } = renderDialog(InsertFunctionDialog);
		fireEvent.change(screen.getByRole('combobox', { name: 'Or select a category:' }), {
			target: { value: 'Text' },
		});
		const list = screen.getByRole('listbox', { name: 'Select a function' });
		fireEvent.click(within(list).getByText('CONCAT'));
		expect(screen.getByTestId('function-syntax').textContent).toMatch(/^CONCAT\(/);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(onClose).toHaveBeenCalled();
		expect(store.state.edit?.text).toBe('=CONCAT(');
		expect(recentFunctions()[0]).toBe('CONCAT');
	});

	it('searches with Enter in the search box without closing', () => {
		const { onClose } = renderDialog(InsertFunctionDialog);
		const search = screen.getByRole('textbox', { name: 'Search for a function:' });
		fireEvent.change(search, { target: { value: 'xlookup' } });
		fireEvent.keyDown(search, { key: 'Enter' });
		expect(onClose).not.toHaveBeenCalled();
		expect(
			(screen.getByRole('combobox', { name: 'Or select a category:' }) as HTMLSelectElement).value,
		).toBe('recommended');
		const list = screen.getByRole('listbox', { name: 'Select a function' });
		expect(within(list).getByRole('option', { selected: true }).textContent).toBe('XLOOKUP');
	});

	it('inserts on double-click and changes nothing on Cancel', () => {
		let view = renderDialog(InsertFunctionDialog);
		fireEvent.doubleClick(within(screen.getByRole('listbox')).getByText('AVERAGE'));
		expect(view.store.state.edit?.text).toBe('=AVERAGE(');
		view.unmount();
		view = renderDialog(InsertFunctionDialog);
		fireEvent.click(screen.getByRole('button', { name: 'Cancel' }));
		expect(view.store.state.edit).toBeNull();
	});
});

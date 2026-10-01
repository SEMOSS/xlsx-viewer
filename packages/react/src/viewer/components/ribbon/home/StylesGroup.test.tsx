import { fireEvent, render, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { ShellContext } from '../../../shell-context';
import { ViewerStoreContext } from '../../../state/use-store';
import { ViewerStore } from '../../../state/viewer-store';
import { makeShell, selectRange } from '../../dialogs/__tests__/render-dialog';
import { StylesGroup } from './StylesGroup';

function setup(readOnly = false) {
	const store = new ViewerStore(createWorkbook(), { readOnly });
	selectRange(store, 'B2:B6');
	const shell = makeShell();
	render(
		<ViewerStoreContext.Provider value={store}>
			<ShellContext.Provider value={shell}>
				<StylesGroup />
			</ShellContext.Provider>
		</ViewerStoreContext.Provider>,
	);
	return { store, shell };
}

function openMenu() {
	fireEvent.click(screen.getByRole('button', { name: 'Conditional Formatting' }));
	return screen.getByRole('menu', { name: 'Conditional Formatting' });
}

function openFlyout(label: string) {
	fireEvent.click(within(openMenu()).getByRole('menuitem', { name: label }));
}

describe('StylesGroup', () => {
	it('lists the Conditional Formatting entries in Excel order', () => {
		setup();
		const menu = openMenu();
		const names = [
			'Highlight Cells Rules',
			'Top/Bottom Rules',
			'Data Bars',
			'Color Scales',
			'Icon Sets',
			'Clear Rules',
			'Manage Rules...',
		];
		const items = within(menu).getAllByRole('menuitem');
		expect(items).toEqual(names.map((name) => within(menu).getByRole('menuitem', { name })));
	});

	it('opens the rule dialogs from the Highlight Cells and Top/Bottom flyouts', () => {
		const { shell } = setup();
		openFlyout('Highlight Cells Rules');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Between...' }));
		expect(shell.openDialog).toHaveBeenCalledWith('conditionalRule', { rule: 'between' });
		openFlyout('Top/Bottom Rules');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Bottom 10%...' }));
		expect(shell.openDialog).toHaveBeenCalledWith('conditionalRule', { rule: 'bottomPercent' });
	});

	it('opens the Rules Manager', () => {
		const { shell } = setup();
		fireEvent.click(within(openMenu()).getByRole('menuitem', { name: 'Manage Rules...' }));
		expect(shell.openDialog).toHaveBeenCalledWith('manageRules');
	});

	it('shows one flyout at a time', () => {
		setup();
		const menu = openMenu();
		fireEvent.pointerEnter(within(menu).getByRole('menuitem', { name: 'Highlight Cells Rules' }));
		expect(screen.getByRole('menu', { name: 'Highlight Cells Rules' })).toBeTruthy();
		fireEvent.pointerEnter(within(menu).getByRole('menuitem', { name: 'Color Scales' }));
		expect(screen.queryByRole('menu', { name: 'Highlight Cells Rules' })).toBeNull();
		expect(screen.getAllByRole('menuitem', { name: /Color Scale$/ })).toHaveLength(12);
	});

	it('applies gallery picks to the selection immediately', () => {
		const { store } = setup();
		openFlyout('Data Bars');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Green Data Bar (Solid Fill)' }));
		openFlyout('Color Scales');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Green - Yellow - Red Color Scale' }));
		openFlyout('Icon Sets');
		fireEvent.click(screen.getByRole('menuitem', { name: '3 Flags' }));
		const rules = store.activeSheet.conditionalFormats.flatMap((f) => f.rules);
		expect(rules.map((r) => r.type).sort()).toEqual(['colorScale', 'dataBar', 'iconSet']);
		expect(rules.find((r) => r.type === 'dataBar')?.dataBar?.gradient).toBe(false);
		expect(rules.find((r) => r.type === 'iconSet')?.iconSet?.iconSet).toBe('3Flags');
		expect(store.activeSheet.conditionalFormats[0].ranges[0]).toEqual({
			top: 1,
			left: 1,
			bottom: 5,
			right: 1,
		});
	});

	it('clears rules from the selection and the sheet', () => {
		const { store } = setup();
		openFlyout('Data Bars');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Blue Data Bar (Gradient Fill)' }));
		openFlyout('Clear Rules');
		fireEvent.click(screen.getByRole('menuitem', { name: 'Clear Rules from Entire Sheet' }));
		expect(store.activeSheet.conditionalFormats).toHaveLength(0);
	});

	it('opens flyouts from the keyboard', () => {
		setup();
		const menu = openMenu();
		expect(document.activeElement).toBe(menu);
		fireEvent.keyDown(menu, { key: 'ArrowDown' });
		expect(document.activeElement).toBe(
			within(menu).getByRole('menuitem', { name: 'Highlight Cells Rules' }),
		);
		const first = within(menu).getByRole('menuitem', { name: 'Highlight Cells Rules' });
		first.focus();
		fireEvent.keyDown(first, { key: 'ArrowRight' });
		expect(screen.getByRole('menu', { name: 'Highlight Cells Rules' })).toBeTruthy();
		expect(document.activeElement).toBe(screen.getByRole('menuitem', { name: 'Greater Than...' }));
	});

	it('opens the Format as Table and Cell Styles galleries', () => {
		setup();
		fireEvent.click(screen.getByRole('button', { name: 'Format as Table' }));
		expect(screen.getByRole('option', { name: 'Table Style Medium 2' })).toBeTruthy();
		fireEvent.keyDown(document.activeElement ?? document.body, { key: 'Escape' });
		fireEvent.click(screen.getByRole('button', { name: 'Cell Styles' }));
		expect(screen.getByRole('option', { name: 'Heading 1' })).toBeTruthy();
	});

	it('disables Conditional Formatting when read-only', () => {
		setup(true);
		expect(
			(screen.getByRole('button', { name: 'Conditional Formatting' }) as HTMLButtonElement)
				.disabled,
		).toBe(true);
	});
});

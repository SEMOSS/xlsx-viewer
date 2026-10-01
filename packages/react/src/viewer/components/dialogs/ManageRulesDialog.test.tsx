import { fireEvent, screen, within } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import {
	addConditionalFormat,
	createCellIsRule,
	createDataBarRule,
	createWorkbook,
	listManagedRules,
	parseRange,
} from 'xlsx-viewer-core';

import {
	formatAppliesTo,
	moveRule,
	parseAppliesTo,
	visibleRules,
	type DraftRule,
} from '../../commands/manage-rules-commands';
import { ViewerStore } from '../../state/viewer-store';
import { renderDialog, selectRange } from './__tests__/render-dialog';
import { ManageRulesDialog } from './ManageRulesDialog';

const r = (a1: string) => parseRange(a1)!;

function workbook() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	addConditionalFormat(sheet, r('A1:A10'), createDataBarRule('FF638EC6'), wb.styles);
	addConditionalFormat(
		sheet,
		r('A1:A10'),
		createCellIsRule('greaterThan', ['5'], 'redText'),
		wb.styles,
	);
	addConditionalFormat(
		sheet,
		r('C1:C5'),
		createCellIsRule('lessThan', ['0'], 'redText'),
		wb.styles,
	);
	return wb;
}

/** Open the manager with `ref` selected. */
function setup(ref = 'A1') {
	const store = new ViewerStore(workbook());
	selectRange(store, ref);
	return renderDialog(ManageRulesDialog, { store });
}

function ruleNames() {
	return within(screen.getByRole('listbox', { name: 'Rules' }))
		.getAllByRole('option')
		.map((o) => o.textContent ?? '');
}

describe('manage-rules helpers', () => {
	it('formats and parses Applies to', () => {
		expect(formatAppliesTo([r('A1:A10'), r('C3')])).toBe('=$A$1:$A$10,$C$3');
		expect(parseAppliesTo('=$A$1:$A$10, C3', 'Sheet1')).toEqual([r('A1:A10'), r('C3')]);
		expect(parseAppliesTo("='Sheet1'!B2:B4", 'Sheet1')).toEqual([r('B2:B4')]);
		expect(parseAppliesTo('=Other!B2', 'Sheet1')).toBeNull();
		expect(parseAppliesTo('=nonsense(', 'Sheet1')).toBeNull();
	});

	it('moves past the next visible rule only', () => {
		const rules: DraftRule[] = listManagedRules(workbook().sheets[0].conditionalFormats).map(
			(m, i) => ({ ...m, id: String(i) }),
		);
		const visible = visibleRules(rules, 'selection', [r('A1')]);
		expect(visible.map((d) => d.id)).toEqual(['1', '2']);
		expect(moveRule(rules, '2', -1, visible).map((d) => d.id)).toEqual(['0', '2', '1']);
	});
});

describe('ManageRulesDialog', () => {
	it('lists rules for the current selection or the whole sheet', () => {
		setup('C2');
		expect(ruleNames()).toEqual(['Cell Value < 0AaBbCcYyZz=$C$1:$C$5']);
		fireEvent.change(screen.getByRole('combobox', { name: 'Show formatting rules for:' }), {
			target: { value: 'sheet' },
		});
		expect(ruleNames().map((n) => n.slice(0, 14))).toEqual([
			'Cell Value < 0',
			'Cell Value > 5',
			'Data Bar=$A$1:',
		]);
	});

	it('reorders, deletes and toggles Stop If True in one undoable step', () => {
		const view = setup();
		const wb = view.store.workbook;
		fireEvent.change(screen.getByRole('combobox', { name: 'Show formatting rules for:' }), {
			target: { value: 'sheet' },
		});
		const list = screen.getByRole('listbox', { name: 'Rules' });
		fireEvent.click(within(list).getAllByRole('option')[1]);
		fireEvent.click(screen.getByRole('button', { name: 'Move Up' }));
		fireEvent.click(screen.getByRole('checkbox', { name: 'Stop If True' }));
		fireEvent.click(within(list).getAllByRole('option')[2]);
		fireEvent.click(screen.getByRole('button', { name: /Delete Rule/ }));
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(view.onClose).toHaveBeenCalled();
		const after = listManagedRules(wb.sheets[0].conditionalFormats);
		expect(after.map((m) => [m.rule.type, m.rule.priority, Boolean(m.rule.stopIfTrue)])).toEqual([
			['cellIs', 1, true],
			['cellIs', 2, false],
		]);
		expect(view.store.getSnapshot().history.undo).toBe('Manage Rules');
	});

	it('edits Applies to and rejects invalid references', () => {
		const view = setup('C2');
		const wb = view.store.workbook;
		const input = screen.getByRole('textbox', { name: 'Applies to:' });
		expect((input as HTMLInputElement).value).toBe('=$C$1:$C$5');
		fireEvent.change(input, { target: { value: '=Q:Q(' } });
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(screen.getByRole('alert').textContent).toBe("Reference isn't valid.");
		expect(view.onClose).not.toHaveBeenCalled();
		fireEvent.change(input, { target: { value: '=$D$1:$D$3' } });
		fireEvent.click(screen.getByRole('button', { name: 'Apply' }));
		expect(listManagedRules(wb.sheets[0].conditionalFormats)[0].ranges).toEqual([r('D1:D3')]);
		expect(view.onClose).not.toHaveBeenCalled();
	});

	it('adds no undo step when nothing changed', () => {
		const view = setup();
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(view.onClose).toHaveBeenCalled();
		expect(view.store.getSnapshot().history.undo).toBeNull();
	});
});

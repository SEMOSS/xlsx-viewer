import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, parseRange } from 'xlsx-viewer-core';

import { createTranslator } from '../../../i18n';
import { CONDITIONAL_RULES, type ConditionalRuleKind } from '../../commands/cf-rule-builder';
import { renderDialog, selectRange, typeCells } from './__tests__/render-dialog';
import { ConditionalRuleDialog } from './ConditionalRuleDialog';

const t = createTranslator();

function setup(rule: ConditionalRuleKind) {
	const wb = createWorkbook();
	typeCells(wb, { A1: '10', A2: '20', A3: '30', A4: '40' });
	const view = renderDialog(ConditionalRuleDialog, { workbook: wb, payload: { rule } });
	selectRange(view.store, 'A1:A4');
	return view;
}

function rules(view: ReturnType<typeof setup>) {
	return view.store.activeSheet.conditionalFormats.flatMap((f) => f.rules);
}

describe('ConditionalRuleDialog', () => {
	it('renders every rule kind with its Excel title and prompt', () => {
		for (const [kind, spec] of Object.entries(CONDITIONAL_RULES)) {
			const view = setup(kind as ConditionalRuleKind);
			expect(screen.getByRole('dialog', { name: t(spec.titleKey) })).toBeTruthy();
			expect(screen.getByText(t(spec.promptKey))).toBeTruthy();
			view.unmount();
		}
	});

	it('creates a Greater Than rule with the chosen format on Enter', () => {
		const view = setup('greaterThan');
		const value = screen.getByRole('textbox', {
			name: t(CONDITIONAL_RULES.greaterThan.promptKey),
		});
		fireEvent.change(value, { target: { value: '25' } });
		fireEvent.change(screen.getByRole('combobox', { name: 'Format with' }), {
			target: { value: 'yellowFillDarkYellowText' },
		});
		fireEvent.keyDown(value, { key: 'Enter' });
		expect(view.onClose).toHaveBeenCalled();
		const [rule] = rules(view);
		expect(rule).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['25'] });
		const dxf = view.store.workbook.styles.dxfs[rule.dxfId!];
		expect(dxf.font?.color).toEqual({ rgb: 'FF9C5700' });
		expect(view.store.activeSheet.conditionalFormats[0].ranges).toEqual([parseRange('A1:A4')]);
	});

	it('creates Between with two values', () => {
		const view = setup('between');
		fireEvent.change(screen.getByRole('textbox', { name: 'From' }), { target: { value: '15' } });
		fireEvent.change(screen.getByRole('textbox', { name: 'To' }), { target: { value: '35' } });
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(rules(view)[0]).toMatchObject({ operator: 'between', formulas: ['15', '35'] });
	});

	it('shows an error and stays open for an invalid rank', () => {
		const view = setup('topPercent');
		fireEvent.change(screen.getByRole('spinbutton'), { target: { value: '500' } });
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(screen.getByRole('alert').textContent).toBe('Enter a whole number between 1 and 100.');
		expect(view.onClose).not.toHaveBeenCalled();
		expect(rules(view)).toHaveLength(0);
	});

	it('creates date, duplicate and average rules', () => {
		let view = setup('dateOccurring');
		fireEvent.change(
			screen.getByRole('combobox', { name: t(CONDITIONAL_RULES.dateOccurring.promptKey) }),
			{ target: { value: 'nextMonth' } },
		);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(rules(view)[0]).toMatchObject({ type: 'timePeriod', timePeriod: 'nextMonth' });
		view.unmount();
		view = setup('duplicateValues');
		fireEvent.change(
			screen.getByRole('combobox', { name: t(CONDITIONAL_RULES.duplicateValues.promptKey) }),
			{ target: { value: 'unique' } },
		);
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(rules(view)[0].type).toBe('uniqueValues');
		view.unmount();
		view = setup('belowAverage');
		fireEvent.click(screen.getByRole('button', { name: 'OK' }));
		expect(rules(view)[0]).toMatchObject({ type: 'aboveAverage', aboveAverage: false });
	});

	it('changes nothing on Cancel', () => {
		const view = setup('textContains');
		fireEvent.click(screen.getByRole('button', { name: 'Cancel' }));
		expect(view.onClose).toHaveBeenCalled();
		expect(rules(view)).toHaveLength(0);
	});
});

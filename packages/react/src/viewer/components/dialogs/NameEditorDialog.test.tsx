import { fireEvent, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import type { DefinedName } from 'xlsx-viewer-core';
import { createWorkbook } from 'xlsx-viewer-core';

import type { NameDraft } from '../../commands/name-commands';
import { renderDialog } from './__tests__/render-dialog';
import { NameEditorDialog } from './NameEditorDialog';
import type { DialogProps } from './types';

const blank: NameDraft = { name: '', scope: undefined, comment: '', refersTo: '=Data!$A$1' };

function setup(
	opts: { initial?: Partial<NameDraft>; editing?: DefinedName; result?: string | null } = {},
) {
	const wb = createWorkbook({ sheetNames: ['Data', 'Other', 'Ghost'] });
	wb.sheets[2].state = 'veryHidden';
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Data!$B$1' },
		{ name: 'Local', formula: '1', localSheet: 1 },
	);
	const onSubmit = vi.fn<(def: DefinedName) => string | null>(() => opts.result ?? null);
	function Editor({ onClose }: DialogProps) {
		return (
			<NameEditorDialog
				title={opts.editing ? 'Edit Name' : 'New Name'}
				initial={{ ...blank, ...opts.initial }}
				editing={opts.editing}
				onSubmit={onSubmit}
				onClose={onClose}
			/>
		);
	}
	const utils = renderDialog(Editor, { workbook: wb });
	return { ...utils, onSubmit };
}

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;
const ok = () => fireEvent.click(screen.getByRole('button', { name: 'OK' }));

describe('NameEditorDialog', () => {
	it('shows the Excel fields and submits a validated definition', () => {
		const { onSubmit } = setup({ initial: { name: 'Tax' } });
		expect(screen.getByRole('dialog', { name: 'New Name' })).toBeTruthy();
		expect(document.activeElement).toBe(field('Name:'));
		const scope = screen.getByRole('combobox', { name: 'Scope' }) as HTMLSelectElement;
		expect([...scope.options].map((o) => o.textContent)).toEqual(['Workbook', 'Data', 'Other']);
		fireEvent.change(scope, { target: { value: '1' } });
		fireEvent.change(field('Comment:'), { target: { value: 'Sales tax' } });
		fireEvent.change(field('Refers to:'), { target: { value: '=$C$3*2' } });
		ok();
		expect(onSubmit).toHaveBeenCalledWith({
			name: 'Tax',
			formula: 'Data!$C$3*2',
			localSheet: 1,
			comment: 'Sales tax',
		});
	});

	it("reports invalid and duplicate names with Excel's messages", () => {
		const { onSubmit } = setup({ initial: { name: 'my rate' } });
		ok();
		expect(screen.getByRole('alert').textContent).toContain(
			'The name that you entered is not valid.',
		);
		expect(screen.getByRole('alert').textContent).toContain(
			'does not begin with a letter or an underscore',
		);
		expect(document.activeElement).toBe(field('Name:'));
		fireEvent.change(field('Name:'), { target: { value: 'rate' } });
		expect(screen.queryByRole('alert')).toBeNull();
		ok();
		expect(screen.getByRole('alert').textContent).toBe(
			'The name already exists. Enter a unique name.',
		);
		fireEvent.change(screen.getByRole('combobox', { name: 'Scope' }), { target: { value: '0' } });
		ok();
		expect(onSubmit).toHaveBeenCalledWith({ name: 'rate', formula: 'Data!$A$1', localSheet: 0 });
	});

	it('rejects a bad formula and turns plain text into a constant', () => {
		const { onSubmit } = setup({ initial: { name: 'Tax' } });
		fireEvent.change(field('Refers to:'), { target: { value: '=SUM(1,' } });
		ok();
		expect(screen.getByRole('alert').textContent).toBe("There's a problem with this formula.");
		expect(document.activeElement).toBe(field('Refers to:'));
		expect(onSubmit).not.toHaveBeenCalled();
		fireEvent.change(field('Refers to:'), { target: { value: 'Hello' } });
		fireEvent.keyDown(field('Refers to:'), { key: 'Enter' });
		expect(onSubmit).toHaveBeenCalledWith({ name: 'Tax', formula: '"Hello"' });
	});

	it('locks the scope when editing and keeps an untouched formula as written', () => {
		const editing: DefinedName = { name: 'Local', formula: '[1]Book!$A$1', localSheet: 1 };
		const { onSubmit } = setup({
			editing,
			initial: { name: 'Local', scope: 1, refersTo: '=[1]Book!$A$1' },
		});
		expect(screen.getByRole('dialog', { name: 'Edit Name' })).toBeTruthy();
		expect((screen.getByRole('combobox', { name: 'Scope' }) as HTMLSelectElement).disabled).toBe(
			true,
		);
		ok();
		expect(onSubmit).toHaveBeenCalledWith({
			name: 'Local',
			formula: '[1]Book!$A$1',
			localSheet: 1,
		});
	});

	it('shows an error returned by onSubmit and keeps Enter in the comment box', () => {
		const { onSubmit, onClose } = setup({ initial: { name: 'Tax' }, result: 'Rejected by host' });
		fireEvent.keyDown(field('Comment:'), { key: 'Enter' });
		expect(onSubmit).not.toHaveBeenCalled();
		fireEvent.keyDown(field('Name:'), { key: 'Enter' });
		expect(screen.getByRole('alert').textContent).toBe('Rejected by host');
		fireEvent.keyDown(field('Name:'), { key: 'Escape' });
		expect(onClose).toHaveBeenCalled();
	});
});

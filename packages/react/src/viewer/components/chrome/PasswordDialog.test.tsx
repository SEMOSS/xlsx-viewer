import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { createWorkbook, putCell, writeWorkbook, XlsxPasswordError } from 'xlsx-viewer-core';

import { ExcelViewer } from '../../ExcelViewer';
import { stubLayout } from './__tests__/layout-stubs';
import { PasswordDialog } from './PasswordDialog';

const INCORRECT = /The password you supplied is not correct\. Verify that the CAPS LOCK key is off/;

let encrypted: Uint8Array;

beforeAll(async () => {
	const wb = createWorkbook({ sheetNames: ['Payroll'] });
	putCell(wb.sheets[0], 0, 0, { value: 'salary' });
	encrypted = await writeWorkbook(wb, { password: 'open sesame' });
});

stubLayout();

const passwordInput = () => screen.getByLabelText('Password:') as HTMLInputElement;

function enter(password: string) {
	fireEvent.change(passwordInput(), { target: { value: password } });
	fireEvent.click(screen.getByRole('button', { name: 'OK' }));
}

describe('opening an encrypted workbook', () => {
	it('asks for the password, reports a wrong one and opens with the right one', async () => {
		const onError = vi.fn();
		render(<ExcelViewer content={encrypted} fileName='Payroll.xlsx' onError={onError} />);
		const dialog = await screen.findByRole('dialog', { name: 'Password' });
		expect(dialog.textContent).toContain("'Payroll.xlsx' is protected.");
		expect(screen.queryByText(INCORRECT)).toBeNull();

		enter('open says me');
		expect(await screen.findByText(INCORRECT)).toBeTruthy();
		expect(passwordInput().value).toBe('');

		enter('open sesame');
		await waitFor(() => expect(screen.queryByRole('dialog', { name: 'Password' })).toBeNull());
		expect(await screen.findByText('Payroll')).toBeTruthy();
		expect(onError).not.toHaveBeenCalled();
	});

	it('shows that the workbook is protected when the prompt is cancelled', async () => {
		const onError = vi.fn();
		render(<ExcelViewer content={encrypted} fileName='Payroll.xlsx' onError={onError} />);
		await screen.findByRole('dialog', { name: 'Password' });
		fireEvent.click(screen.getByRole('button', { name: 'Cancel' }));
		const alert = await screen.findByRole('alert');
		expect(alert.textContent).toContain('This workbook is protected with a password.');
		expect(onError).toHaveBeenCalledWith(expect.any(XlsxPasswordError));
		// The prompt can be opened again.
		fireEvent.click(screen.getByRole('button', { name: 'Enter Password...' }));
		expect(await screen.findByRole('dialog', { name: 'Password' })).toBeTruthy();
	});

	it('opens without a prompt when the host passes the password', async () => {
		render(<ExcelViewer content={encrypted} fileName='Payroll.xlsx' password='open sesame' />);
		expect(await screen.findByText('Payroll')).toBeTruthy();
		expect(screen.queryByRole('dialog', { name: 'Password' })).toBeNull();
	});
});

describe('PasswordDialog', () => {
	it('submits with Enter and cancels with Escape', () => {
		const onSubmit = vi.fn();
		const onCancel = vi.fn();
		render(
			<PasswordDialog
				fileName='a.xlsx'
				incorrect={false}
				onSubmit={onSubmit}
				onCancel={onCancel}
			/>,
		);
		fireEvent.change(passwordInput(), { target: { value: 'pw' } });
		fireEvent.keyDown(passwordInput(), { key: 'Enter' });
		expect(onSubmit).toHaveBeenCalledWith('pw');
		fireEvent.keyDown(passwordInput(), { key: 'Escape' });
		expect(onCancel).toHaveBeenCalled();
		expect(passwordInput().type).toBe('password');
	});
});

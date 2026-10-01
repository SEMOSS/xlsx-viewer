import { act, fireEvent, render, screen } from '@testing-library/react';
import { createRef } from 'react';
import { describe, expect, it, vi } from 'vitest';
import {
	createWorkbook,
	isEncryptedPackageFile,
	readWorkbook,
	writeWorkbook,
	XlsxPasswordError,
} from 'xlsx-viewer-core';

import { exportWorkbook } from '../../download';
import { ExcelViewer } from '../../ExcelViewer';
import type { ExcelViewerHandle } from '../../types';
import { stubLayout } from './__tests__/layout-stubs';
import { EncryptDialog } from './EncryptDialog';

stubLayout();

const field = (label: string) => screen.getByLabelText(label) as HTMLInputElement;
const type = (label: string, value: string) =>
	fireEvent.change(field(label), { target: { value } });
const ok = () => fireEvent.click(screen.getByRole('button', { name: 'OK' }));

describe('EncryptDialog', () => {
	it('requires the confirmation to match', () => {
		const onSubmit = vi.fn();
		render(<EncryptDialog password={undefined} onSubmit={onSubmit} onClose={vi.fn()} />);
		expect(screen.getByRole('dialog', { name: 'Encrypt Document' })).toBeTruthy();
		type('Password:', 'abc');
		type('Reenter password:', 'abd');
		ok();
		expect(screen.getByRole('alert').textContent).toBe('The passwords do not match.');
		expect(onSubmit).not.toHaveBeenCalled();
		type('Reenter password:', 'abc');
		expect(screen.queryByRole('alert')).toBeNull();
		fireEvent.keyDown(field('Reenter password:'), { key: 'Enter' });
		expect(onSubmit).toHaveBeenCalledWith('abc');
	});

	it('keeps the current password, and removes it when both fields are cleared', () => {
		const onSubmit = vi.fn();
		render(<EncryptDialog password='old' onSubmit={onSubmit} onClose={vi.fn()} />);
		expect(field('Password:').value).toBe('old');
		ok();
		expect(onSubmit).toHaveBeenLastCalledWith('old');
		type('Password:', '');
		type('Reenter password:', '');
		ok();
		expect(onSubmit).toHaveBeenLastCalledWith(undefined);
	});
});

describe('File > Info > Protect Workbook', () => {
	async function openInfo() {
		fireEvent.click(screen.getByRole('button', { name: 'File' }));
		return screen.findByText('Protect Workbook');
	}

	it('encrypts saves with the password and marks the workbook changed', async () => {
		const plain = await writeWorkbook(createWorkbook({ sheetNames: ['Budget'] }));
		const ref = createRef<ExcelViewerHandle>();
		const onSave = vi.fn();
		const onDirtyChange = vi.fn();
		render(
			<ExcelViewer
				ref={ref}
				content={plain}
				fileName='Budget.xlsx'
				onSave={onSave}
				onDirtyChange={onDirtyChange}
			/>,
		);
		await screen.findByText('Budget');
		expect(onDirtyChange).not.toHaveBeenCalled();

		await openInfo();
		expect(screen.getByText('Require a password to open this workbook.')).toBeTruthy();
		fireEvent.click(screen.getByRole('button', { name: 'Encrypt with Password...' }));
		type('Password:', 'Tr0ub4dor');
		type('Reenter password:', 'Tr0ub4dor');
		ok();
		expect(screen.getByText('A password is required to open this workbook.')).toBeTruthy();
		expect(onDirtyChange).toHaveBeenLastCalledWith(true);

		const content = await ref.current!.getContent();
		expect(isEncryptedPackageFile(content)).toBe(true);
		await expect(readWorkbook(content)).rejects.toBeInstanceOf(XlsxPasswordError);
		expect((await readWorkbook(content, { password: 'Tr0ub4dor' })).sheets[0].name).toBe('Budget');
		const exported = await exportWorkbook(ref.current!.getWorkbook()!, 0, 'xlsx');
		expect(isEncryptedPackageFile(exported as Uint8Array)).toBe(true);
		await act(() => ref.current!.save());
		expect(isEncryptedPackageFile(onSave.mock.calls[0][0])).toBe(true);
		expect(onDirtyChange).toHaveBeenLastCalledWith(false);

		// Removing the password saves plain files again.
		fireEvent.click(screen.getByRole('button', { name: 'Encrypt with Password...' }));
		type('Password:', '');
		type('Reenter password:', '');
		ok();
		expect(screen.getByText('Require a password to open this workbook.')).toBeTruthy();
		expect(onDirtyChange).toHaveBeenLastCalledWith(true);
		expect(isEncryptedPackageFile(await ref.current!.getContent())).toBe(false);
	});

	it('keeps the password an encrypted file was opened with', async () => {
		const locked = await writeWorkbook(createWorkbook({ sheetNames: ['Secret'] }), {
			password: 'kept',
		});
		const ref = createRef<ExcelViewerHandle>();
		render(<ExcelViewer ref={ref} content={locked} fileName='Secret.xlsx' password='kept' />);
		await screen.findByText('Secret');
		await openInfo();
		expect(screen.getByText('A password is required to open this workbook.')).toBeTruthy();
		const content = await ref.current!.getContent();
		expect((await readWorkbook(content, { password: 'kept' })).encrypted).toBe(true);
	});
});

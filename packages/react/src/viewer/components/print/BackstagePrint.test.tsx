import { fireEvent, render, screen, within } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, putCell } from 'xlsx-viewer-core';

import { ShellContext, type ShellActions } from '../../shell-context';
import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { Backstage } from '../chrome/Backstage';

function setup() {
	const wb = createWorkbook();
	putCell(wb.sheets[0], 0, 0, { value: 'Hello' });
	const store = new ViewerStore(wb);
	const shell: ShellActions = {
		openDialog: vi.fn(),
		save: vi.fn(),
		download: vi.fn(),
		print: vi.fn(),
		notify: vi.fn(),
		themeMode: 'light',
	};
	const onClose = vi.fn();
	render(
		<ViewerStoreContext.Provider value={store}>
			<ShellContext.Provider value={shell}>
				<Backstage fileName='Book.xlsx' onClose={onClose} />
			</ShellContext.Provider>
		</ViewerStoreContext.Provider>,
	);
	return { shell, onClose };
}

describe('File > Print', () => {
	it('opens the print view inside the backstage and goes back to Info', () => {
		const { shell, onClose } = setup();
		const nav = screen.getByRole('navigation');
		fireEvent.click(within(nav).getByRole('button', { name: /Print/ }));
		expect(onClose).not.toHaveBeenCalled();
		expect(shell.print).not.toHaveBeenCalled();
		expect(screen.getByRole('img', { name: 'Page 1' })).toBeTruthy();
		expect(screen.getByText('of 1')).toBeTruthy();
		fireEvent.click(within(nav).getByRole('button', { name: /Info/ }));
		expect(screen.queryByRole('img', { name: 'Page 1' })).toBeNull();
	});
});

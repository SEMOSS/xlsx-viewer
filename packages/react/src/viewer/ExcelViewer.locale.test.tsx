import { act, fireEvent, render, screen } from '@testing-library/react';
import { createRef } from 'react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { de } from '../i18n';
import { stubLayout } from './components/chrome/__tests__/layout-stubs';
import { ExcelViewer } from './ExcelViewer';
import type { ExcelViewerHandle } from './types';

stubLayout();

/** The hidden textarea that receives grid shortcuts. */
function focusSink(): HTMLTextAreaElement {
	const sink = document.querySelector<HTMLTextAreaElement>('.xv-focus-sink');
	if (!sink) {
		throw new Error('the grid did not render');
	}
	return sink;
}

describe('ExcelViewer locale', () => {
	it('renders the ribbon and dialogs in German', async () => {
		render(<ExcelViewer workbook={createWorkbook()} locale='de' />);
		expect(await screen.findByRole('tab', { name: 'Einfügen' })).toBeTruthy();
		expect(screen.getByRole('button', { name: 'Datei' })).toBeTruthy();
		expect(screen.queryByRole('tab', { name: 'Insert' })).toBeNull();

		fireEvent.keyDown(focusSink(), { key: '1', code: 'Digit1', ctrlKey: true });
		expect(await screen.findByRole('dialog', { name: 'Zellen formatieren' })).toBeTruthy();
	});

	it('writes chrome numbers the way the locale does', async () => {
		render(<ExcelViewer workbook={createWorkbook()} locale='de' />);
		// Testing Library collapses the no-break space of '100\u00a0%' when matching.
		const zoom = await screen.findByText('100 %');
		expect(zoom.textContent).toBe('100\u00a0%');
	});

	it('shows command status messages through the store translator', async () => {
		render(<ExcelViewer workbook={createWorkbook()} locale='de' />);
		fireEvent.click(await screen.findByRole('tab', { name: de['tab.formulas'] }));
		fireEvent.click(screen.getByRole('button', { name: de['cmd.traceDependents'] }));
		expect((await screen.findByRole('status')).textContent).toBe(de['message.noDependents']);
	});

	it('switches language when the locale prop changes, keeping the workbook and its state', async () => {
		const workbook = createWorkbook({ sheetNames: ['Budget'] });
		const ref = createRef<ExcelViewerHandle>();
		const { rerender } = render(<ExcelViewer ref={ref} workbook={workbook} locale='en' />);
		expect(await screen.findByRole('tab', { name: 'Insert' })).toBeTruthy();
		act(() => {
			ref.current?.select('C3');
		});

		rerender(<ExcelViewer ref={ref} workbook={workbook} locale='fr' />);
		expect(await screen.findByRole('tab', { name: 'Insertion' })).toBeTruthy();
		expect(screen.queryByRole('tab', { name: 'Insert' })).toBeNull();
		// Same store: the selection made before the switch is still there.
		expect(screen.getByDisplayValue('C3')).toBeTruthy();
		expect(ref.current?.getWorkbook()).toBe(workbook);

		rerender(<ExcelViewer ref={ref} workbook={workbook} locale='zh-CN' />);
		expect(await screen.findByRole('tab', { name: '插入' })).toBeTruthy();
		expect(screen.getByDisplayValue('C3')).toBeTruthy();
	});

	it('applies translation overrides on top of the locale', async () => {
		render(
			<ExcelViewer
				workbook={createWorkbook()}
				locale='de'
				translations={{ 'tab.insert': 'Hinzufügen' }}
			/>,
		);
		expect(await screen.findByRole('tab', { name: 'Hinzufügen' })).toBeTruthy();
		expect(screen.getByRole('tab', { name: 'Start' })).toBeTruthy();
	});
});

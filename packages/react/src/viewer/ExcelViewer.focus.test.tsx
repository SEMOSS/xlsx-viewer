import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { stubLayout } from './components/chrome/__tests__/layout-stubs';
import { ExcelViewer } from './ExcelViewer';

stubLayout();

async function renderViewer() {
	render(<ExcelViewer workbook={createWorkbook({ sheetNames: ['Budget', 'Notes'] })} />);
	await screen.findByRole('tab', { name: 'Insert' });
	const sink = document.querySelector<HTMLTextAreaElement>('.xv-focus-sink');
	if (!sink) {
		throw new Error('the grid did not render');
	}
	return sink;
}

/** A mouse press as the browser sends it: cancelable, so the viewer can keep focus. */
function press(el: Element): MouseEvent {
	const e = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 });
	el.dispatchEvent(e);
	return e;
}

describe('ExcelViewer keyboard focus', () => {
	it('keeps the grid focused when a ribbon button is clicked, so typing edits the cell', async () => {
		const sink = await renderViewer();
		sink.focus();
		const bold = screen.getByRole('button', { name: 'Bold' });
		expect(press(bold).defaultPrevented).toBe(true);
		fireEvent.click(bold);
		expect(document.activeElement).toBe(sink);
	});

	it('returns focus to the grid from a focused toolbar button', async () => {
		const sink = await renderViewer();
		const bold = screen.getByRole('button', { name: 'Bold' });
		bold.focus();
		press(screen.getByRole('button', { name: 'Italic' }));
		expect(document.activeElement).toBe(sink);
	});

	it('returns focus to the grid after Enter in the Name Box or the font size box', async () => {
		const sink = await renderViewer();
		const nameBox = screen.getByDisplayValue('A1');
		nameBox.focus();
		fireEvent.change(nameBox, { target: { value: 'C3' } });
		fireEvent.keyDown(nameBox, { key: 'Enter' });
		expect(document.activeElement).toBe(sink);

		const size = screen.getByRole('textbox', { name: 'Font Size' });
		size.focus();
		fireEvent.keyDown(size, { key: 'Escape' });
		expect(document.activeElement).toBe(sink);
	});

	it('returns focus to the grid when a sheet tab is clicked', async () => {
		const sink = await renderViewer();
		const notes = screen.getByRole('tab', { name: 'Notes' });
		notes.focus();
		fireEvent.click(notes);
		expect(notes.getAttribute('aria-selected')).toBe('true');
		expect(document.activeElement).toBe(sink);
	});

	it('renames a sheet on Enter, cancels on Escape, and returns focus to the grid', async () => {
		const sink = await renderViewer();
		const rename = (key: string, value: string) => {
			fireEvent.doubleClick(screen.getByRole('tab', { name: 'Notes' }));
			const input = screen.getByRole('textbox', { name: 'Rename' });
			expect(document.activeElement).toBe(input);
			fireEvent.change(input, { target: { value } });
			fireEvent.keyDown(input, { key });
			expect(screen.queryByRole('textbox', { name: 'Rename' })).toBeNull();
			expect(document.activeElement).toBe(sink);
		};
		rename('Escape', 'Ignored');
		expect(screen.getByRole('tab', { name: 'Notes' })).toBeTruthy();
		expect(screen.queryByRole('tab', { name: 'Ignored' })).toBeNull();
		rename('Enter', 'Summary');
		expect(screen.getByRole('tab', { name: 'Summary' })).toBeTruthy();
		expect(screen.queryByRole('tab', { name: 'Notes' })).toBeNull();
	});

	it('gives the backstage the keyboard while open and the grid afterwards', async () => {
		const sink = await renderViewer();
		sink.focus();
		fireEvent.click(screen.getByRole('button', { name: 'File' }));
		const back = await screen.findByRole('button', { name: 'Back' });
		expect(document.activeElement).toBe(back);
		fireEvent.keyDown(back, { key: 'Escape' });
		expect(screen.queryByRole('button', { name: 'Back' })).toBeNull();
		expect(document.activeElement).toBe(sink);
	});
});

import { fireEvent, render, screen } from '@testing-library/react';
import { useRef, useState } from 'react';
import { describe, expect, it } from 'vitest';

import { Popover } from './Popover';

function Harness() {
	const [open, setOpen] = useState(false);
	const anchor = useRef<HTMLButtonElement>(null);
	return (
		<>
			<textarea className='xv-focus-sink' aria-label='Spreadsheet' />
			<button ref={anchor} type='button' onClick={() => setOpen(true)}>
				Fill Color
			</button>
			<Popover open={open} onClose={() => setOpen(false)} anchorRef={anchor} ariaLabel='Colors'>
				<input aria-label='Hex' />
				<button type='button' onClick={() => setOpen(false)}>
					Apply
				</button>
			</Popover>
		</>
	);
}

function openPopover() {
	render(<Harness />);
	const sink = screen.getByLabelText('Spreadsheet');
	sink.focus();
	fireEvent.click(screen.getByText('Fill Color'));
	return sink;
}

describe('Popover focus', () => {
	it('gives focus back to where it was when closing with focus inside', () => {
		const sink = openPopover();
		screen.getByLabelText('Hex').focus();
		fireEvent.click(screen.getByText('Apply'));
		expect(screen.queryByRole('dialog', { name: 'Colors' })).toBeNull();
		expect(document.activeElement).toBe(sink);
	});

	it('leaves focus alone when it already moved elsewhere', () => {
		openPopover();
		const anchor = screen.getByText('Fill Color');
		anchor.focus();
		fireEvent.keyDown(anchor, { key: 'Escape' });
		expect(screen.queryByRole('dialog', { name: 'Colors' })).toBeNull();
		expect(document.activeElement).toBe(anchor);
	});

	it('closes when the user types into the grid', () => {
		const sink = openPopover();
		fireEvent.keyDown(sink, { key: 'Shift' });
		expect(screen.getByRole('dialog', { name: 'Colors' })).toBeTruthy();
		fireEvent.keyDown(sink, { key: 'a' });
		expect(screen.queryByRole('dialog', { name: 'Colors' })).toBeNull();
		expect(document.activeElement).toBe(sink);
	});
});

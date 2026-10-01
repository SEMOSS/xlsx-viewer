import { ChevronDown } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

import { useT } from '../../../i18n';
import { useReturnToGrid } from '../../shell-context';
import { Popover } from './Popover';

interface ComboBoxProps {
	value: string;
	options: readonly string[];
	onCommit: (value: string) => void;
	label: string;
	width: number;
	/** Render each option in its own font (font picker). */
	previewFont?: boolean;
	disabled?: boolean;
}

/** Editable dropdown (Font name / Font size). Enter or picking an option commits. */
export function ComboBox({
	value,
	options,
	onCommit,
	label,
	width,
	previewFont,
	disabled,
}: ComboBoxProps) {
	const t = useT();
	const returnToGrid = useReturnToGrid();
	const anchor = useRef<HTMLDivElement>(null);
	const input = useRef<HTMLInputElement>(null);
	const [text, setText] = useState(value);
	const [open, setOpen] = useState(false);
	const [focused, setFocused] = useState(false);
	useEffect(() => {
		if (!focused) {
			setText(value);
		}
	}, [value, focused]);
	const commit = (v: string) => {
		setOpen(false);
		if (v.trim() && v !== value) {
			onCommit(v.trim());
		}
	};
	return (
		<div
			ref={anchor}
			className='xv-combo inline-flex h-7 shrink-0 items-center rounded-sm border border-input bg-card text-[12px] focus-within:border-primary'
			style={{ width }}
		>
			<input
				ref={input}
				aria-label={label}
				title={label}
				disabled={disabled}
				value={text}
				className='h-full min-w-0 flex-1 bg-transparent px-1.5 text-foreground outline-none'
				onFocus={(e) => {
					setFocused(true);
					e.currentTarget.select();
				}}
				onBlur={() => setFocused(false)}
				onChange={(e) => setText(e.target.value)}
				onKeyDown={(e) => {
					if (e.key === 'Enter') {
						e.preventDefault();
						commit(text);
						returnToGrid(input.current);
					} else if (e.key === 'Escape') {
						setText(value);
						returnToGrid(input.current);
					} else if (e.key === 'ArrowDown') {
						setOpen(true);
					}
				}}
			/>
			<button
				type='button'
				tabIndex={-1}
				aria-label={t('ui.comboBoxList', { label })}
				disabled={disabled}
				className='inline-flex h-full w-5 items-center justify-center text-muted-foreground hover:text-foreground'
				onClick={() => setOpen((v) => !v)}
			>
				<ChevronDown size={14} />
			</button>
			<Popover
				open={open}
				onClose={() => setOpen(false)}
				anchorRef={anchor}
				role='listbox'
				ariaLabel={label}
			>
				<div className='max-h-[320px] overflow-y-auto py-1' style={{ minWidth: width }}>
					{options.map((o) => (
						<button
							key={o}
							type='button'
							role='option'
							aria-selected={o === value}
							className={`block w-full px-2.5 py-1 text-left text-[13px] hover:bg-accent ${o === value ? 'bg-accent font-medium' : ''}`}
							style={previewFont ? { fontFamily: `'${o}', sans-serif` } : undefined}
							onClick={() => commit(o)}
						>
							{o}
						</button>
					))}
				</div>
			</Popover>
		</div>
	);
}

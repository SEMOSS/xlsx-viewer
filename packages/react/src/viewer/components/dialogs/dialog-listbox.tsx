import { useEffect, useId, useRef, type KeyboardEvent, type ReactNode } from 'react';

import { cx } from './dialog-cx';

export interface ListBoxProps<T> {
	items: readonly T[];
	getKey: (item: T) => string;
	render: (item: T, selected: boolean) => ReactNode;
	/** Key of the selected item. */
	value: string | null;
	onChange: (key: string) => void;
	/** Double-click on an item. */
	onActivate?: (key: string) => void;
	/** Text used for type-ahead (defaults to none). */
	getText?: (item: T) => string;
	label: string;
	className?: string;
	/** Rendered when the list is empty. */
	empty?: ReactNode;
}

/**
 * Single-selection list box (`role="listbox"`): arrows, Home / End and
 * Page Up / Down move the selection, letters jump by type-ahead. Enter is
 * left to the enclosing dialog (OK).
 */
export function ListBox<T>({
	items,
	getKey,
	render,
	value,
	onChange,
	onActivate,
	getText,
	label,
	className,
	empty,
}: ListBoxProps<T>) {
	const id = useId();
	const listRef = useRef<HTMLDivElement>(null);
	const typed = useRef({ text: '', at: 0 });
	const index = items.findIndex((it) => getKey(it) === value);
	const optionId = (i: number) => `${id}-opt-${i}`;

	useEffect(() => {
		if (index < 0) {
			return;
		}
		const el = listRef.current?.querySelector<HTMLElement>(`[id="${optionId(index)}"]`);
		el?.scrollIntoView?.({ block: 'nearest' });
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [index]);

	const move = (to: number) => {
		if (items.length === 0) {
			return;
		}
		const i = Math.max(0, Math.min(items.length - 1, to));
		onChange(getKey(items[i]));
	};

	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		const page = 8;
		switch (e.key) {
			case 'ArrowDown':
				move(index + 1);
				break;
			case 'ArrowUp':
				move(index < 0 ? 0 : index - 1);
				break;
			case 'Home':
				move(0);
				break;
			case 'End':
				move(items.length - 1);
				break;
			case 'PageDown':
				move(index + page);
				break;
			case 'PageUp':
				move(index - page);
				break;
			default: {
				if (!getText || e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey) {
					return;
				}
				const now = Date.now();
				const t = typed.current;
				t.text = now - t.at > 700 ? e.key.toLowerCase() : t.text + e.key.toLowerCase();
				t.at = now;
				const hit = items.findIndex((it) => getText(it).toLowerCase().startsWith(t.text));
				if (hit >= 0) {
					move(hit);
				}
				break;
			}
		}
		if (e.key !== 'Enter' && e.key !== 'Tab' && e.key !== 'Escape') {
			e.preventDefault();
		}
	};

	return (
		<div
			ref={listRef}
			role='listbox'
			tabIndex={0}
			aria-label={label}
			aria-activedescendant={index >= 0 ? optionId(index) : undefined}
			className={cx(
				'overflow-y-auto rounded-md border border-input bg-card py-0.5 text-[13px] outline-none focus:border-primary',
				className,
			)}
			onKeyDown={onKeyDown}
		>
			{items.length === 0 && empty}
			{items.map((it, i) => {
				const key = getKey(it);
				const selected = key === value;
				return (
					<div
						key={key}
						id={optionId(i)}
						role='option'
						aria-selected={selected}
						className={cx(
							'cursor-default select-none px-2 py-[3px]',
							selected ? 'bg-primary text-primary-foreground' : 'hover:bg-accent',
						)}
						onClick={() => {
							listRef.current?.focus();
							onChange(key);
						}}
						onDoubleClick={() => onActivate?.(key)}
					>
						{render(it, selected)}
					</div>
				);
			})}
		</div>
	);
}

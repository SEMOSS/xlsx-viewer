import { ChevronRight } from 'lucide-react';
import { useEffect, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';

import { Popover } from '../../ui/Popover';

export interface CascadeEntry {
	id: string;
	label?: string;
	icon?: ReactNode;
	separator?: boolean;
	disabled?: boolean;
	onSelect?: () => void;
	/** Flyout content (a list or a gallery); `close` closes the whole menu. */
	submenu?: (close: () => void) => ReactNode;
}

const ITEM =
	'xv-menu-item flex w-full items-center gap-2.5 px-3 py-1.5 text-left text-[13px] outline-none hover:bg-accent focus-visible:bg-accent disabled:pointer-events-none disabled:opacity-40';

/** Move focus between the buttons of a menu with the arrow keys, Home and End. */
export function menuArrowKeys(
	e: KeyboardEvent<HTMLElement>,
	selector = 'button:not(:disabled)',
): boolean {
	if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
		return false;
	}
	const items = [...e.currentTarget.querySelectorAll<HTMLElement>(selector)];
	if (items.length === 0) {
		return false;
	}
	const i = items.indexOf(document.activeElement as HTMLElement);
	const next =
		e.key === 'Home'
			? 0
			: e.key === 'End'
				? items.length - 1
				: e.key === 'ArrowDown'
					? (i + 1) % items.length
					: (i - 1 + items.length) % items.length;
	e.preventDefault();
	items[next].focus();
	return true;
}

/**
 * A dropdown menu with flyout submenus (only one open at a time). Hover or
 * Right / Enter opens a flyout, Left / Escape closes it, arrows move.
 */
export function CascadeMenu({
	entries,
	close,
	label,
}: {
	entries: readonly CascadeEntry[];
	close: () => void;
	label: string;
}) {
	const [open, setOpen] = useState<string | null>(null);
	const rows = useRef(new Map<string, HTMLButtonElement | null>());
	const root = useRef<HTMLDivElement>(null);
	const flyout = useRef<HTMLDivElement>(null);
	const focusFlyout = useRef(false);
	// Stable anchor ref for the flyout (a new object per render would re-place it endlessly).
	const anchor = useRef<HTMLElement | null>(null);
	const openEntry = entries.find((e) => e.id === open && e.submenu);
	anchor.current = openEntry ? (rows.current.get(openEntry.id) ?? null) : null;

	// Take focus when the menu opens so the arrow keys work straight away.
	useEffect(() => {
		root.current?.focus({ preventScroll: true });
	}, []);

	useEffect(() => {
		if (open && focusFlyout.current) {
			focusFlyout.current = false;
			flyout.current?.querySelector<HTMLElement>('button:not(:disabled)')?.focus();
		}
	}, [open]);

	const openFlyout = (id: string, byKey: boolean) => {
		focusFlyout.current = byKey;
		setOpen(id);
	};

	const onFlyoutKey = (e: KeyboardEvent<HTMLDivElement>) => {
		// Portal events bubble to the root menu in React; keep flyout navigation local.
		e.stopPropagation();
		if (e.key === 'ArrowLeft') {
			e.preventDefault();
			const id = open;
			setOpen(null);
			if (id) {
				rows.current.get(id)?.focus();
			}
			return;
		}
		menuArrowKeys(e);
	};

	return (
		<div
			ref={root}
			role='menu'
			tabIndex={-1}
			aria-label={label}
			className='flex min-w-[260px] flex-col py-1 outline-none'
			onKeyDown={(e) => menuArrowKeys(e, ':scope > button:not(:disabled)')}
		>
			{entries.map((entry) => {
				if (entry.separator) {
					return <div key={entry.id} role='separator' className='my-1 h-px bg-border' />;
				}
				const hasSub = Boolean(entry.submenu);
				return (
					<button
						key={entry.id}
						ref={(el) => {
							rows.current.set(entry.id, el);
						}}
						type='button'
						role='menuitem'
						aria-haspopup={hasSub ? 'menu' : undefined}
						aria-expanded={hasSub ? open === entry.id : undefined}
						disabled={entry.disabled}
						className={`${ITEM} ${open === entry.id ? 'bg-accent' : ''}`}
						onPointerEnter={() => setOpen(hasSub ? entry.id : null)}
						onClick={() => {
							if (hasSub) {
								openFlyout(entry.id, false);
							} else {
								close();
								entry.onSelect?.();
							}
						}}
						onKeyDown={(e) => {
							if (hasSub && (e.key === 'ArrowRight' || e.key === 'Enter' || e.key === ' ')) {
								e.preventDefault();
								e.stopPropagation();
								openFlyout(entry.id, true);
							}
						}}
					>
						<span
							aria-hidden='true'
							className='flex w-5 shrink-0 justify-center text-muted-foreground'
						>
							{entry.icon}
						</span>
						<span className='flex-1 whitespace-nowrap'>{entry.label}</span>
						{hasSub && <ChevronRight size={14} className='text-muted-foreground' />}
					</button>
				);
			})}
			{openEntry?.submenu && (
				<Popover
					open
					onClose={() => setOpen(null)}
					anchorRef={anchor}
					placement='right-start'
					role='presentation'
				>
					<div ref={flyout} onKeyDown={onFlyoutKey}>
						{openEntry.submenu(close)}
					</div>
				</Popover>
			)}
		</div>
	);
}

export interface FlyoutItem {
	id: string;
	label: string;
	icon?: ReactNode;
	onSelect: () => void;
}

/** A plain list flyout (Highlight Cells Rules, Clear Rules, ...). */
export function FlyoutList({
	items,
	close,
	label,
}: {
	items: readonly FlyoutItem[];
	close: () => void;
	label: string;
}) {
	return (
		<div role='menu' aria-label={label} className='flex min-w-[230px] flex-col py-1'>
			{items.map((item) => (
				<button
					key={item.id}
					type='button'
					role='menuitem'
					className={ITEM}
					onClick={() => {
						close();
						item.onSelect();
					}}
				>
					<span
						aria-hidden='true'
						className='flex w-5 shrink-0 justify-center text-muted-foreground'
					>
						{item.icon}
					</span>
					<span className='flex-1 whitespace-nowrap'>{item.label}</span>
				</button>
			))}
		</div>
	);
}

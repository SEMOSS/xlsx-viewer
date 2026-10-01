import { ChevronRight, Check } from 'lucide-react';
import { useRef, useState, type ReactNode } from 'react';

import { Popover } from './Popover';

export interface MenuItem {
	id?: string;
	label?: string;
	icon?: ReactNode;
	shortcut?: string;
	disabled?: boolean;
	checked?: boolean;
	danger?: boolean;
	onSelect?: () => void;
	/** Nested items, shown on hover. */
	items?: MenuItem[];
	/** Custom content (e.g. a color grid) instead of a clickable row. */
	render?: (close: () => void) => ReactNode;
	separator?: boolean;
	/** Non-interactive section heading. */
	heading?: string;
	/** Tooltip text. */
	title?: string;
}

interface MenuListProps {
	items: MenuItem[];
	onClose: () => void;
	className?: string;
}

function SubmenuRow({
	item,
	onClose,
	open,
	onOpen,
}: {
	item: MenuItem;
	onClose: () => void;
	open: boolean;
	onOpen: (open: boolean) => void;
}) {
	const ref = useRef<HTMLButtonElement>(null);
	return (
		<>
			<button
				ref={ref}
				type='button'
				role='menuitem'
				aria-haspopup='menu'
				aria-expanded={open}
				disabled={item.disabled}
				className='xv-menu-item flex w-full items-center gap-2.5 px-3 py-1.5 text-left text-[13px] hover:bg-accent disabled:opacity-40'
				onPointerEnter={() => onOpen(true)}
				onClick={() => onOpen(true)}
			>
				<span className='flex w-4 shrink-0 justify-center text-muted-foreground'>{item.icon}</span>
				<span className='flex-1'>{item.label}</span>
				<ChevronRight size={14} className='text-muted-foreground' />
			</button>
			<Popover
				open={open}
				onClose={() => onOpen(false)}
				anchorRef={ref}
				placement='right-start'
				role='menu'
			>
				<MenuList items={item.items ?? []} onClose={onClose} />
			</Popover>
		</>
	);
}

/** Vertical list of menu items. */
export function MenuList({ items, onClose, className }: MenuListProps) {
	// One submenu open at a time: hovering another row closes it.
	const [openSub, setOpenSub] = useState<string | null>(null);
	return (
		<div role='menu' className={`flex min-w-[200px] flex-col py-1 ${className ?? ''}`}>
			{items.map((item, i) => {
				const key = item.id ?? `${item.label ?? 'sep'}-${i}`;
				if (item.separator) {
					return <div key={key} role='separator' className='my-1 h-px bg-border' />;
				}
				if (item.heading) {
					return (
						<div
							key={key}
							className='px-3 pt-2 pb-1 text-[11px] font-semibold uppercase tracking-wide text-muted-foreground'
						>
							{item.heading}
						</div>
					);
				}
				if (item.render) {
					return <div key={key}>{item.render(onClose)}</div>;
				}
				if (item.items) {
					return (
						<SubmenuRow
							key={key}
							item={item}
							onClose={onClose}
							open={openSub === key}
							onOpen={(open) => setOpenSub((cur) => (open ? key : cur === key ? null : cur))}
						/>
					);
				}
				return (
					<button
						key={key}
						type='button'
						role={item.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'}
						aria-checked={item.checked}
						disabled={item.disabled}
						title={item.title}
						className={`xv-menu-item flex w-full items-center gap-2.5 px-3 py-1.5 text-left text-[13px] hover:bg-accent disabled:pointer-events-none disabled:opacity-40 ${item.danger ? 'text-destructive' : ''}`}
						onPointerEnter={() => setOpenSub(null)}
						onClick={() => {
							onClose();
							item.onSelect?.();
						}}
					>
						<span className='flex w-4 shrink-0 justify-center text-muted-foreground'>
							{item.checked ? <Check size={14} /> : item.icon}
						</span>
						<span className='flex-1 whitespace-nowrap'>{item.label}</span>
						{item.shortcut && (
							<span className='pl-4 text-[11px] text-muted-foreground'>{item.shortcut}</span>
						)}
					</button>
				);
			})}
		</div>
	);
}

/** A button-anchored dropdown menu. */
export function DropdownMenu({
	anchorRef,
	open,
	onClose,
	items,
	placement,
}: {
	anchorRef: React.RefObject<HTMLElement | null>;
	open: boolean;
	onClose: () => void;
	items: MenuItem[];
	placement?: 'bottom-start' | 'bottom-end';
}) {
	return (
		<Popover open={open} onClose={onClose} anchorRef={anchorRef} placement={placement} role='menu'>
			<MenuList items={items} onClose={onClose} />
		</Popover>
	);
}

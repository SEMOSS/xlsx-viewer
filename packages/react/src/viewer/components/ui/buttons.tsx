import { ChevronDown } from 'lucide-react';
import { forwardRef, useRef, useState, type ButtonHTMLAttributes, type ReactNode } from 'react';

import { useT } from '../../../i18n';
import { DropdownMenu, type MenuItem } from './Menu';
import { Popover } from './Popover';

function cx(...parts: Array<string | false | undefined | null>): string {
	return parts.filter(Boolean).join(' ');
}

export interface RibbonButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
	icon?: ReactNode;
	label?: string;
	/** Visible text next to the icon. */
	showLabel?: boolean;
	pressed?: boolean;
	shortcut?: string;
}

/** Compact ribbon / toolbar button with a tooltip. */
export const RibbonButton = forwardRef<HTMLButtonElement, RibbonButtonProps>(function RibbonButton(
	{ icon, label, showLabel, pressed, shortcut, className, ...rest },
	ref,
) {
	const t = useT();
	const title = label ? (shortcut ? t('ui.withShortcut', { label, shortcut }) : label) : undefined;
	return (
		<button
			ref={ref}
			type='button'
			title={title}
			aria-label={label}
			aria-pressed={pressed}
			className={cx(
				'xv-rbtn inline-flex h-7 min-w-7 shrink-0 items-center justify-center gap-1.5 rounded-sm px-1.5 text-[12px] text-foreground transition-colors hover:bg-accent disabled:pointer-events-none disabled:opacity-40',
				pressed && 'bg-accent text-primary shadow-[inset_0_0_0_1px_var(--xlsx-primary)]',
				className,
			)}
			{...rest}
		>
			{icon}
			{showLabel && label && <span className='whitespace-nowrap'>{label}</span>}
		</button>
	);
});

interface SplitButtonProps {
	icon?: ReactNode;
	label: string;
	showLabel?: boolean;
	onClick?: () => void;
	disabled?: boolean;
	pressed?: boolean;
	/** Menu items for the dropdown half. */
	items?: MenuItem[];
	/** Custom dropdown content instead of items. */
	renderPanel?: (close: () => void) => ReactNode;
	shortcut?: string;
}

/** A button with a separate dropdown arrow (Excel's split buttons). */
export function SplitButton({
	icon,
	label,
	showLabel,
	onClick,
	disabled,
	pressed,
	items,
	renderPanel,
	shortcut,
}: SplitButtonProps) {
	const t = useT();
	const anchor = useRef<HTMLDivElement>(null);
	const [open, setOpen] = useState(false);
	const close = () => setOpen(false);
	return (
		<div
			ref={anchor}
			className='xv-split inline-flex h-7 shrink-0 items-stretch rounded-sm hover:shadow-[inset_0_0_0_1px_var(--xlsx-border)]'
		>
			<RibbonButton
				icon={icon}
				label={label}
				showLabel={showLabel}
				pressed={pressed}
				disabled={disabled}
				shortcut={shortcut}
				onClick={onClick ?? (() => setOpen(true))}
				className='rounded-r-none'
			/>
			<button
				type='button'
				aria-label={t('ui.splitButtonOptions', { label })}
				aria-haspopup='menu'
				aria-expanded={open}
				disabled={disabled}
				className='inline-flex w-4 items-center justify-center rounded-r-sm text-muted-foreground hover:bg-accent disabled:opacity-40'
				onClick={() => setOpen((v) => !v)}
			>
				<ChevronDown size={12} />
			</button>
			{items ? (
				<DropdownMenu anchorRef={anchor} open={open} onClose={close} items={items} />
			) : (
				<Popover open={open} onClose={close} anchorRef={anchor} ariaLabel={label}>
					{renderPanel?.(close)}
				</Popover>
			)}
		</div>
	);
}

interface MenuButtonProps {
	icon?: ReactNode;
	label: string;
	showLabel?: boolean;
	items?: MenuItem[];
	renderPanel?: (close: () => void) => ReactNode;
	disabled?: boolean;
	className?: string;
}

/** A button that only opens a dropdown (e.g. Conditional Formatting, Insert). */
export function MenuButton({
	icon,
	label,
	showLabel = true,
	items,
	renderPanel,
	disabled,
	className,
}: MenuButtonProps) {
	const ref = useRef<HTMLButtonElement>(null);
	const [open, setOpen] = useState(false);
	const close = () => setOpen(false);
	return (
		<>
			<button
				ref={ref}
				type='button'
				title={label}
				aria-label={label}
				aria-haspopup='menu'
				aria-expanded={open}
				disabled={disabled}
				className={cx(
					'xv-rbtn inline-flex h-7 shrink-0 items-center gap-1 rounded-sm px-1.5 text-[12px] text-foreground hover:bg-accent disabled:pointer-events-none disabled:opacity-40',
					open && 'bg-accent',
					className,
				)}
				onClick={() => setOpen((v) => !v)}
			>
				{icon}
				{showLabel && <span className='whitespace-nowrap'>{label}</span>}
				<ChevronDown size={12} className='text-muted-foreground' />
			</button>
			{items ? (
				<DropdownMenu anchorRef={ref} open={open} onClose={close} items={items} />
			) : (
				<Popover open={open} onClose={close} anchorRef={ref} ariaLabel={label}>
					{renderPanel?.(close)}
				</Popover>
			)}
		</>
	);
}

/** Visual group of joined controls (shared muted background, like the reference ribbon). */
export function ControlCluster({ children }: { children: ReactNode }) {
	return (
		<div className='xv-cluster inline-flex items-center rounded-sm bg-muted/70 p-0'>{children}</div>
	);
}

export { cx };

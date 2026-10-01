import { X } from 'lucide-react';
import { useEffect, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

import { useT } from '../../../i18n';
import { usePortalHost } from './Popover';

interface DialogProps {
	title: string;
	onClose: () => void;
	children: ReactNode;
	footer?: ReactNode;
	width?: number;
	/** Element to focus when the dialog opens (defaults to the first input). */
	initialFocusSelector?: string;
}

/** Modal dialog with a title bar, body and optional footer. */
export function Dialog({
	title,
	onClose,
	children,
	footer,
	width = 420,
	initialFocusSelector,
}: DialogProps) {
	const t = useT();
	const host = usePortalHost();
	const ref = useRef<HTMLDivElement>(null);
	const prevFocus = useRef<Element | null>(null);

	useEffect(() => {
		prevFocus.current = document.activeElement;
		const el = ref.current;
		const target = el?.querySelector<HTMLElement>(
			initialFocusSelector ?? 'input, select, textarea, button[data-primary]',
		);
		// Without a control to focus, the dialog itself takes the keyboard (not the grid behind it).
		(target ?? el)?.focus();
		if (target instanceof HTMLInputElement) {
			target.select();
		}
		return () => {
			(prevFocus.current as HTMLElement | null)?.focus?.({ preventScroll: true });
		};
	}, [initialFocusSelector]);

	const node = (
		<div
			data-xv-modal
			className='fixed inset-0 z-[900] flex items-center justify-center bg-black/25'
			onPointerDown={(e) => {
				if (e.target === e.currentTarget) {
					onClose();
				}
			}}
		>
			<div
				ref={ref}
				role='dialog'
				aria-modal='true'
				aria-label={title}
				tabIndex={-1}
				className='xv-dialog flex max-h-[90vh] flex-col rounded-lg outline-none border border-border bg-popover text-popover-foreground shadow-2xl'
				style={{ width }}
				onKeyDown={(e) => {
					if (e.key === 'Escape') {
						e.stopPropagation();
						onClose();
					}
				}}
			>
				<div className='flex items-center justify-between border-b border-border px-4 py-2.5'>
					<h2 className='text-[14px] font-semibold'>{title}</h2>
					<button
						type='button'
						aria-label={t('dialog.close')}
						className='rounded-sm p-1 text-muted-foreground hover:bg-accent'
						onClick={onClose}
					>
						<X size={16} />
					</button>
				</div>
				<div className='min-h-0 flex-1 overflow-y-auto px-4 py-3 text-[13px]'>{children}</div>
				{footer && (
					<div className='flex justify-end gap-2 border-t border-border px-4 py-2.5'>{footer}</div>
				)}
			</div>
		</div>
	);
	return host ? createPortal(node, host) : node;
}

export function DialogButton({
	children,
	primary,
	onClick,
	disabled,
	type = 'button',
}: {
	children: ReactNode;
	primary?: boolean;
	onClick?: () => void;
	disabled?: boolean;
	type?: 'button' | 'submit';
}) {
	return (
		<button
			type={type}
			data-primary={primary || undefined}
			disabled={disabled}
			onClick={onClick}
			className={
				primary
					? 'min-w-[76px] rounded-md bg-primary px-3 py-1.5 text-[13px] font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50'
					: 'min-w-[76px] rounded-md border border-border bg-card px-3 py-1.5 text-[13px] hover:bg-accent disabled:opacity-50'
			}
		>
			{children}
		</button>
	);
}

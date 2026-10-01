import { type KeyboardEvent, type ReactNode } from 'react';

import { useT } from '../../../i18n';
import { Dialog, DialogButton } from '../ui/Dialog';

export { FIELD, FIELD_SM } from './dialog-fields';
export * from './dialog-tabs';
export * from './dialog-listbox';
export * from './dialog-controls';

/**
 * Shared building blocks for the feature dialogs: a dialog frame where Enter
 * means OK, tabs, list boxes, check boxes, radio buttons and group boxes,
 * all styled like the rest of the viewer chrome.
 */

/** Elements that handle Enter themselves (buttons, multi-line text, links). */
function handlesEnter(target: HTMLElement): boolean {
	if (target.closest('[data-enter="self"]')) {
		return true;
	}
	const tag = target.tagName;
	return tag === 'TEXTAREA' || tag === 'BUTTON' || tag === 'A';
}

export interface FormDialogProps {
	title: string;
	onClose: () => void;
	/** OK / Enter. */
	onOk: () => void;
	/** Caption of the OK button (default: the translated "OK"). */
	okLabel?: string;
	okDisabled?: boolean;
	/** Caption of the Cancel button (default: the translated "Cancel"). */
	cancelLabel?: string;
	width?: number;
	/** Extra buttons shown before Cancel / OK (e.g. Apply). */
	extraButtons?: ReactNode;
	/** Buttons at the far left of the footer (e.g. Clear All, Remove Link). */
	leftButtons?: ReactNode;
	initialFocusSelector?: string;
	children: ReactNode;
}

/**
 * Modal dialog with Cancel / OK buttons. Enter anywhere in the body runs
 * OK (except on buttons, in text areas or inside `data-enter="self"`), and
 * Escape cancels (handled by `Dialog`).
 */
export function FormDialog({
	title,
	onClose,
	onOk,
	okLabel,
	okDisabled,
	cancelLabel,
	width,
	extraButtons,
	leftButtons,
	initialFocusSelector,
	children,
}: FormDialogProps) {
	const t = useT();
	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		if (
			e.key !== 'Enter' ||
			e.shiftKey ||
			e.altKey ||
			e.ctrlKey ||
			e.metaKey ||
			e.nativeEvent.isComposing ||
			handlesEnter(e.target as HTMLElement)
		) {
			return;
		}
		e.preventDefault();
		if (!okDisabled) {
			onOk();
		}
	};
	return (
		<Dialog
			title={title}
			onClose={onClose}
			width={width}
			initialFocusSelector={initialFocusSelector}
			footer={
				<>
					{leftButtons && <div className='mr-auto flex gap-2'>{leftButtons}</div>}
					{extraButtons}
					<DialogButton onClick={onClose}>{cancelLabel ?? t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={onOk} disabled={okDisabled}>
						{okLabel ?? t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<div className='xv-form-dialog flex flex-col gap-3' onKeyDown={onKeyDown}>
				{children}
			</div>
		</Dialog>
	);
}

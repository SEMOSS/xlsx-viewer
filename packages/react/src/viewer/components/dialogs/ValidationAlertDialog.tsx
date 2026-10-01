import { DEFAULT_VALIDATION_ERROR_MESSAGE, DEFAULT_VALIDATION_ERROR_TITLE } from 'xlsx-viewer-core';
import { CircleAlert, CircleX, Info } from 'lucide-react';

import { useT } from '../../../i18n';
import { cancelEdit, commitEdit } from '../../commands/edit-commands';
import type { ValidationAlertPayload } from '../../commands/validation-commands';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';
import type { DialogProps } from './types';

/**
 * Data validation error alert, as Excel shows it for an invalid entry:
 * Stop (Retry / Cancel), Warning ("Continue?" Yes / No / Cancel) or
 * Information (OK / Cancel).
 */
export function ValidationAlertDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const p = (payload ?? {}) as Partial<ValidationAlertPayload> & { fillSelection?: boolean };
	const style = p.style ?? 'stop';
	// The rule's own title and message are workbook content; the defaults are viewer text.
	const message =
		p.message === DEFAULT_VALIDATION_ERROR_MESSAGE
			? t('dialog.validationAlert.defaultMessage')
			: p.message;
	const accept = () => {
		onClose();
		commitEdit(store, {
			move: p.move ?? 'down',
			fillSelection: p.fillSelection,
			skipValidation: true,
		});
	};
	const retry = () => onClose();
	const cancel = () => {
		onClose();
		cancelEdit(store);
	};
	const icon =
		style === 'stop' ? (
			<CircleX size={30} className='shrink-0 text-destructive' />
		) : style === 'warning' ? (
			<CircleAlert size={30} className='shrink-0 text-amber-500' />
		) : (
			<Info size={30} className='shrink-0 text-primary' />
		);
	const buttons =
		style === 'stop' ? (
			<>
				<DialogButton primary onClick={retry}>
					{t('validation.retry')}
				</DialogButton>
				<DialogButton onClick={cancel}>{t('dialog.cancel')}</DialogButton>
			</>
		) : style === 'warning' ? (
			<>
				<DialogButton primary onClick={accept}>
					{t('validation.yes')}
				</DialogButton>
				<DialogButton onClick={retry}>{t('validation.no')}</DialogButton>
				<DialogButton onClick={cancel}>{t('dialog.cancel')}</DialogButton>
			</>
		) : (
			<>
				<DialogButton primary onClick={accept}>
					{t('dialog.ok')}
				</DialogButton>
				<DialogButton onClick={cancel}>{t('dialog.cancel')}</DialogButton>
			</>
		);
	return (
		<Dialog
			// Excel's default alert title names Excel itself; the viewer shows its own name instead.
			title={p.title && p.title !== DEFAULT_VALIDATION_ERROR_TITLE ? p.title : t('app.name')}
			onClose={retry}
			width={420}
			initialFocusSelector='button[data-primary]'
			footer={buttons}
		>
			<div className='flex items-start gap-3'>
				{icon}
				<div role='alert' className='whitespace-pre-wrap leading-relaxed'>
					{message}
					{style === 'warning' && <p className='mt-2'>{t('validation.continue')}</p>}
				</div>
			</div>
		</Dialog>
	);
}

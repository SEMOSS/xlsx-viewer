import { useState } from 'react';

import { useT } from '../../../i18n';
import { Dialog, DialogButton } from '../ui/Dialog';
import { passwordField } from './PasswordDialog';

interface EncryptDialogProps {
	/** The current password (prefilled, so OK keeps it). */
	password: string | undefined;
	/** The new password; `undefined` when both fields were cleared (removes encryption). */
	onSubmit: (password: string | undefined) => void;
	onClose: () => void;
}

/** Excel's "Encrypt Document" dialog: a password and its confirmation. Blank removes it. */
export function EncryptDialog({ password: current, onSubmit, onClose }: EncryptDialogProps) {
	const t = useT();
	const [password, setPassword] = useState(current ?? '');
	const [confirm, setConfirm] = useState(current ?? '');
	const [mismatch, setMismatch] = useState(false);
	const apply = () => {
		if (password !== confirm) {
			setMismatch(true);
			return;
		}
		onSubmit(password || undefined);
	};
	const input = (value: string, set: (v: string) => void) => (
		<input
			type='password'
			className={passwordField}
			value={value}
			maxLength={255}
			autoComplete='new-password'
			onChange={(e) => {
				set(e.target.value);
				setMismatch(false);
			}}
			onKeyDown={(e) => e.key === 'Enter' && apply()}
		/>
	);
	return (
		<Dialog
			title={t('password.encryptTitle')}
			onClose={onClose}
			width={420}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={apply}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<div className='flex flex-col gap-2.5'>
				<p className='font-medium'>{t('password.encryptText')}</p>
				<label className='flex flex-col gap-1'>
					<span>{t('password.label')}</span>
					{input(password, setPassword)}
				</label>
				<label className='flex flex-col gap-1'>
					<span>{t('password.confirmLabel')}</span>
					{input(confirm, setConfirm)}
				</label>
				{mismatch && (
					<p role='alert' className='text-[12px] text-destructive'>
						{t('password.mismatch')}
					</p>
				)}
				<p className='text-[12px] text-muted-foreground'>{t('password.caution')}</p>
			</div>
		</Dialog>
	);
}

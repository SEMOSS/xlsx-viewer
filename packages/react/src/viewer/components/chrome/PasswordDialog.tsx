import { useState } from 'react';

import { useT } from '../../../i18n';
import { Dialog, DialogButton } from '../ui/Dialog';

export const passwordField =
	'w-full rounded-md border border-input bg-card px-2 py-1.5 text-[13px] text-foreground outline-none focus:border-primary';

interface PasswordDialogProps {
	fileName: string;
	/** The last password tried was wrong. */
	incorrect: boolean;
	onSubmit: (password: string) => void;
	onCancel: () => void;
}

/** Excel's "Password" prompt for opening an encrypted workbook. */
export function PasswordDialog({ fileName, incorrect, onSubmit, onCancel }: PasswordDialogProps) {
	const t = useT();
	const [password, setPassword] = useState('');
	const submit = () => onSubmit(password);
	return (
		<Dialog
			title={t('password.title')}
			onClose={onCancel}
			width={380}
			footer={
				<>
					<DialogButton onClick={onCancel}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={submit}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<div className='flex flex-col gap-2'>
				<p>{t('password.protected', { name: fileName })}</p>
				<label className='flex flex-col gap-1'>
					<span>{t('password.label')}</span>
					<input
						type='password'
						className={passwordField}
						value={password}
						maxLength={255}
						autoComplete='off'
						onChange={(e) => setPassword(e.target.value)}
						onKeyDown={(e) => e.key === 'Enter' && submit()}
					/>
				</label>
				{incorrect && (
					<p role='alert' className='text-[12px] text-destructive'>
						{t('password.incorrect')}
					</p>
				)}
			</div>
		</Dialog>
	);
}

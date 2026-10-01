import { Lock } from 'lucide-react';
import { useState } from 'react';

import { useT } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { EncryptDialog } from './EncryptDialog';
import { useWorkbookPassword } from './password-context';

/** File > Info > Protect Workbook: Encrypt with Password, and whether opening needs one. */
export function ProtectWorkbook() {
	const t = useT();
	const store = useViewerStore();
	const state = useWorkbookPassword();
	const [encrypting, setEncrypting] = useState(false);
	if (!state) {
		return null;
	}
	const locked = state.password !== undefined;
	const tone = locked ? 'border-amber-500/60 bg-amber-500/10' : 'border-border bg-card';
	return (
		<div className={`flex items-start gap-3 rounded-md border p-3 ${tone}`}>
			<Lock
				size={22}
				aria-hidden
				className={`mt-0.5 shrink-0 ${locked ? 'text-amber-600' : 'text-primary'}`}
			/>
			<div className='flex min-w-0 flex-col items-start gap-1'>
				<h3 className='text-[14px] font-semibold'>{t('password.protectWorkbook')}</h3>
				<p className='text-[12.5px] text-muted-foreground'>
					{locked ? t('password.requiredToOpen') : t('password.encryptDetail')}
				</p>
				<button
					type='button'
					disabled={store.readOnly}
					onClick={() => setEncrypting(true)}
					className='mt-1 rounded-md border border-border bg-card px-3 py-1.5 text-[13px] hover:bg-accent disabled:opacity-50'
				>
					{t('password.encryptWithPassword')}
				</button>
			</div>
			{encrypting && (
				<EncryptDialog
					password={state.password}
					onClose={() => setEncrypting(false)}
					onSubmit={(password) => {
						state.setPassword(password);
						setEncrypting(false);
					}}
				/>
			)}
		</div>
	);
}

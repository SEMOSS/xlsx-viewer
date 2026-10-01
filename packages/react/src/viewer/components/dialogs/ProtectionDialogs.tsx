import type { ProtectableAction } from 'xlsx-viewer-core';
import { useState } from 'react';

import { useT, type TranslationKey } from '../../../i18n';
import {
	protectSheetCmd,
	protectWorkbookCmd,
	unprotectCmd,
} from '../../commands/protection-commands';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';
import type { DialogProps } from './types';

const field =
	'w-full rounded-md border border-input bg-card px-2 py-1.5 text-[13px] text-foreground outline-none focus:border-primary';

/** A message box with OK (protected-sheet warnings and similar). */
export function AlertDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const message = (payload as { message?: string; title?: string } | undefined)?.message ?? '';
	const title = (payload as { title?: string } | undefined)?.title ?? t('app.name');
	return (
		<Dialog
			title={title}
			onClose={onClose}
			width={440}
			initialFocusSelector='button[data-primary]'
			footer={
				<DialogButton primary onClick={onClose}>
					{t('dialog.ok')}
				</DialogButton>
			}
		>
			<p role='alert' className='leading-relaxed'>
				{message}
			</p>
		</Dialog>
	);
}

/** Excel's "Allow all users of this worksheet to" list, with OOXML flags (true = locked). */
const ALLOW: Array<{ action: ProtectableAction; key: TranslationKey; allowedByDefault: boolean }> =
	[
		{ action: 'selectLockedCells', key: 'protect.selectLocked', allowedByDefault: true },
		{ action: 'selectUnlockedCells', key: 'protect.selectUnlocked', allowedByDefault: true },
		{ action: 'formatCells', key: 'protect.formatCells', allowedByDefault: false },
		{ action: 'formatColumns', key: 'protect.formatColumns', allowedByDefault: false },
		{ action: 'formatRows', key: 'protect.formatRows', allowedByDefault: false },
		{ action: 'insertColumns', key: 'protect.insertColumns', allowedByDefault: false },
		{ action: 'insertRows', key: 'protect.insertRows', allowedByDefault: false },
		{ action: 'insertHyperlinks', key: 'protect.insertHyperlinks', allowedByDefault: false },
		{ action: 'deleteColumns', key: 'protect.deleteColumns', allowedByDefault: false },
		{ action: 'deleteRows', key: 'protect.deleteRows', allowedByDefault: false },
		{ action: 'sort', key: 'protect.sort', allowedByDefault: false },
		{ action: 'autoFilter', key: 'protect.autoFilter', allowedByDefault: false },
		{ action: 'pivotTables', key: 'protect.pivotTables', allowedByDefault: false },
		{ action: 'objects', key: 'protect.objects', allowedByDefault: false },
		{ action: 'scenarios', key: 'protect.scenarios', allowedByDefault: false },
	];

/** Password plus "Reenter password to proceed", as Excel asks when a password is set. */
function usePasswordPair() {
	const [password, setPassword] = useState('');
	const [confirm, setConfirm] = useState('');
	const [confirming, setConfirming] = useState(false);
	return { password, setPassword, confirm, setConfirm, confirming, setConfirming };
}

/** Review > Protect Sheet. */
export function ProtectSheetDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const pw = usePasswordPair();
	const [allowed, setAllowed] = useState<Record<string, boolean>>(() =>
		Object.fromEntries(ALLOW.map((a) => [a.action, a.allowedByDefault])),
	);
	const [busy, setBusy] = useState(false);
	const [error, setError] = useState<string | null>(null);

	const submit = async () => {
		if (pw.password && !pw.confirming) {
			pw.setConfirming(true);
			return;
		}
		if (pw.password && pw.confirm !== pw.password) {
			setError(t('protect.mismatch'));
			pw.setConfirm('');
			return;
		}
		setBusy(true);
		const flags = Object.fromEntries(ALLOW.map((a) => [a.action, !allowed[a.action]]));
		await protectSheetCmd(store, pw.password, flags);
		onClose();
	};

	return (
		<Dialog
			title={pw.confirming ? t('protect.confirmTitle') : t('protect.sheetTitle')}
			onClose={onClose}
			width={380}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary disabled={busy} onClick={() => void submit()}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			{pw.confirming ? (
				<label className='flex flex-col gap-1'>
					<span>{t('protect.reenter')}</span>
					<input
						type='password'
						className={field}
						value={pw.confirm}
						onChange={(e) => pw.setConfirm(e.target.value)}
						onKeyDown={(e) => e.key === 'Enter' && void submit()}
						autoFocus
					/>
					<span className='text-[12px] text-muted-foreground'>{t('protect.caution')}</span>
				</label>
			) : (
				<div className='flex flex-col gap-2'>
					<label className='flex flex-col gap-1'>
						<span>{t('protect.passwordToUnprotect')}</span>
						<input
							type='password'
							className={field}
							value={pw.password}
							onChange={(e) => pw.setPassword(e.target.value)}
							onKeyDown={(e) => e.key === 'Enter' && void submit()}
						/>
					</label>
					<div>
						<div className='mb-1'>{t('protect.allowAll')}</div>
						<div className='max-h-[220px] overflow-y-auto rounded-md border border-border px-2 py-1'>
							{ALLOW.map((a) => (
								<label key={a.action} className='flex items-center gap-2 py-0.5'>
									<input
										type='checkbox'
										checked={allowed[a.action]}
										onChange={(e) => setAllowed((s) => ({ ...s, [a.action]: e.target.checked }))}
									/>
									{t(a.key)}
								</label>
							))}
						</div>
					</div>
				</div>
			)}
			{error && (
				<p role='alert' className='mt-2 text-[12px] text-destructive'>
					{error}
				</p>
			)}
		</Dialog>
	);
}

/** Review > Protect Workbook (structure). */
export function ProtectWorkbookDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const pw = usePasswordPair();
	const [error, setError] = useState<string | null>(null);
	const [busy, setBusy] = useState(false);
	const submit = async () => {
		if (pw.password && !pw.confirming) {
			pw.setConfirming(true);
			return;
		}
		if (pw.password && pw.confirm !== pw.password) {
			setError(t('protect.mismatch'));
			pw.setConfirm('');
			return;
		}
		setBusy(true);
		await protectWorkbookCmd(store, pw.password);
		onClose();
	};
	return (
		<Dialog
			title={pw.confirming ? t('protect.confirmTitle') : t('protect.workbookTitle')}
			onClose={onClose}
			width={360}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary disabled={busy} onClick={() => void submit()}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			{pw.confirming ? (
				<label className='flex flex-col gap-1'>
					<span>{t('protect.reenter')}</span>
					<input
						type='password'
						className={field}
						value={pw.confirm}
						onChange={(e) => pw.setConfirm(e.target.value)}
						onKeyDown={(e) => e.key === 'Enter' && void submit()}
						autoFocus
					/>
				</label>
			) : (
				<div className='flex flex-col gap-2'>
					<p>{t('protect.workbookStructure')}</p>
					<label className='flex flex-col gap-1'>
						<span>{t('protect.passwordOptional')}</span>
						<input
							type='password'
							className={field}
							value={pw.password}
							onChange={(e) => pw.setPassword(e.target.value)}
							onKeyDown={(e) => e.key === 'Enter' && void submit()}
						/>
					</label>
				</div>
			)}
			{error && (
				<p role='alert' className='mt-2 text-[12px] text-destructive'>
					{error}
				</p>
			)}
		</Dialog>
	);
}

/** Unprotect Sheet / Unprotect Workbook password prompt; payload `{ target: 'sheet' | 'workbook' }`. */
export function UnprotectDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const target = (payload as { target?: 'sheet' | 'workbook' } | undefined)?.target ?? 'sheet';
	const [password, setPassword] = useState('');
	const [error, setError] = useState<string | null>(null);
	const [busy, setBusy] = useState(false);
	const submit = async () => {
		setBusy(true);
		const ok = await unprotectCmd(store, target, password);
		setBusy(false);
		if (ok) {
			onClose();
		} else {
			setError(t('protect.wrongPassword'));
			setPassword('');
		}
	};
	return (
		<Dialog
			title={target === 'sheet' ? t('protect.unprotectSheet') : t('protect.unprotectWorkbook')}
			onClose={onClose}
			width={340}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary disabled={busy} onClick={() => void submit()}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<label className='flex flex-col gap-1'>
				<span>{t('protect.password')}</span>
				<input
					type='password'
					className={field}
					value={password}
					onChange={(e) => setPassword(e.target.value)}
					onKeyDown={(e) => e.key === 'Enter' && void submit()}
				/>
			</label>
			{error && (
				<p role='alert' className='mt-2 text-[12px] text-destructive'>
					{error}
				</p>
			)}
		</Dialog>
	);
}
